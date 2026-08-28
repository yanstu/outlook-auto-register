#!/usr/bin/env bash
# 部署 outlook-auto-register 到 38 服务器：
#   cd outlook-auto-register && ./scripts/deploy_to_38.sh
#
# 凭据放在 outlook-auto-register/.deploy/local.env（0600，已 gitignore），脚本启动时
# 自动读取，不需要每次手输主机/密码。模板见 .deploy/local.env.example。
#
#   OUTLOOK_DEPLOY_HOST            SSH 目标，形如 root@38.76.160.89
#   OUTLOOK_DEPLOY_PASSWORD        SSH 密码；非空则用 sshpass，否则用 ~/.ssh 密钥
#   OUTLOOK_DOMAIN                 域名，默认 outlook.coolhs.com
#   OUTLOOK_CONSOLE_PASSWORD       运维台登录口令；不能为空（服务挂在公网域名上）
#   OUTLOOK_MAILBOX_API_ADMIN_KEY  Mailbox API 的 admin key；留空则远端沿用已有的，
#                                  远端也没有时现场生成一枚
#   OUTLOOK_SKIP_NGINX=1           只更新应用，不动 nginx / 证书
#
# 流程：rsync 源码 → 建 venv 装依赖 → 渲染远端 .env → 装 systemd --user 服务
# （outlook-console，端口 8890）→ 配 nginx + certbot → 本地与域名探活。
#
# 应用始终跑在远端 ubuntu 用户下（WorkingDirectory=/home/ubuntu/outlook-auto-register）。
# SSH 以 root 登录时，脚本会把 venv / systemd --user 相关动作降权到 ubuntu 执行，
# 避免在 /home/ubuntu 下留 root 属主的文件、或把 user 服务装到 root 名下。
#
# 口令与 key 全程不回显：本地只打掩码，远端只写进 0600 的 .env 与 systemd 单元。
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"   # …/outlook-auto-register
REMOTE_DIR="/home/ubuntu/outlook-auto-register"
APP_USER="ubuntu"
PORT=8890

LOCAL_ENV="${ROOT}/.deploy/local.env"
if [[ -f "${LOCAL_ENV}" ]]; then
  echo "==> 读取凭据 ${LOCAL_ENV#"${ROOT}/"}"
  set -a
  # shellcheck disable=SC1090
  source "${LOCAL_ENV}"
  set +a
fi

DOMAIN="${OUTLOOK_DOMAIN:-outlook.coolhs.com}"
REMOTE="${OUTLOOK_DEPLOY_HOST:-root@38.76.160.89}"
SKIP_NGINX="${OUTLOOK_SKIP_NGINX:-0}"
CONSOLE_PASSWORD="${OUTLOOK_CONSOLE_PASSWORD:-}"
ADMIN_KEY="${OUTLOOK_MAILBOX_API_ADMIN_KEY:-}"

# 运维台没有账号体系，一旦上公网，空口令等于把整个账号库摆在门口。
if [[ -z "${CONSOLE_PASSWORD}" ]]; then
  echo "OUTLOOK_CONSOLE_PASSWORD 不能为空（服务将暴露在 https://${DOMAIN}）" >&2
  echo "把它写进 .deploy/local.env，或临时 export 后重跑。" >&2
  exit 1
fi

mask() { local s="$1"; if [[ ${#s} -le 4 ]]; then printf '****'; else printf '%s****%s' "${s:0:2}" "${s: -2}"; fi; }

SSH_OPTS=(
  -o StrictHostKeyChecking=no
  -o UserKnownHostsFile=/dev/null
  -o LogLevel=ERROR
  -o ConnectTimeout=20
  -o ServerAliveInterval=15
)

if [[ -n "${OUTLOOK_DEPLOY_PASSWORD:-}" ]]; then
  if ! command -v sshpass >/dev/null 2>&1; then
    echo "==> 安装 sshpass"
    sudo apt-get update -qq && sudo apt-get install -y -qq sshpass
  fi
  export SSHPASS="${OUTLOOK_DEPLOY_PASSWORD}"
  SSH=(sshpass -e ssh "${SSH_OPTS[@]}")
  RSH="sshpass -e ssh ${SSH_OPTS[*]}"
  AUTH_KIND="密码 (sshpass)"
else
  SSH=(ssh "${SSH_OPTS[@]}")
  RSH="ssh ${SSH_OPTS[*]}"
  AUTH_KIND="SSH 密钥"
fi

echo "==> 目标 ${REMOTE}:${REMOTE_DIR}  端口 ${PORT}  域名 ${DOMAIN}"
echo "==> 认证 ${AUTH_KIND}  应用用户 ${APP_USER}"
echo "==> 运维台口令 $(mask "${CONSOLE_PASSWORD}")（完整值只在 .deploy/local.env 与远端 systemd 单元里）"

echo "==> rsync → ${REMOTE}:${REMOTE_DIR}"
# accounts/ 排除：线上账号库（含 refresh token）与导出文件都在那儿，
# --delete 一来就会连库带号一起清掉。.env 同理：远端那份由本脚本单独渲染，
# 里面有 captcha key / 代理 / Mailbox API admin key，不能被本机的开发配置覆盖。
rsync -rlptz --delete -e "${RSH}" \
  --exclude '.venv/' \
  --exclude 'accounts/' \
  --exclude '.git/' \
  --exclude '__pycache__/' \
  --exclude '*.pyc' \
  --exclude '.env' \
  --exclude '.deploy/' \
  --exclude '.pytest_cache/' \
  --exclude 'node_modules/' \
  --exclude '*.db' \
  "${ROOT}/" "${REMOTE}:${REMOTE_DIR}/"

echo "==> 远端安装 + systemd + nginx"
{
  # 变量用 %q 转义后注入，远端脚本本体走 quoted heredoc，免去逐个 \$ 转义。
  # 走 stdin 而不是命令行参数，口令不会出现在远端的 ps 里。
  printf 'REMOTE_DIR=%q\n'        "${REMOTE_DIR}"
  printf 'PORT=%q\n'              "${PORT}"
  printf 'DOMAIN=%q\n'            "${DOMAIN}"
  printf 'APP_USER=%q\n'          "${APP_USER}"
  printf 'SKIP_NGINX=%q\n'        "${SKIP_NGINX}"
  printf 'CONSOLE_PASSWORD=%q\n'  "${CONSOLE_PASSWORD}"
  printf 'ADMIN_KEY=%q\n'         "${ADMIN_KEY}"
  cat <<'REMOTE_SCRIPT'
set -euo pipefail

APP_UID="$(id -u "${APP_USER}")"
APP_HOME="$(getent passwd "${APP_USER}" | cut -d: -f6)"
ENV_FILE="${REMOTE_DIR}/.env"

if [[ "$(id -u)" -eq 0 ]]; then
  SUDO=()
else
  SUDO=(sudo)
fi

if [[ "$(id -un)" == "${APP_USER}" ]]; then
  AS_APP=(env)
elif command -v sudo >/dev/null 2>&1; then
  AS_APP=(sudo -u "${APP_USER}" env)
elif [[ "$(id -u)" -eq 0 ]] && command -v runuser >/dev/null 2>&1; then
  AS_APP=(runuser -u "${APP_USER}" -- env)
else
  echo "无法切换到 ${APP_USER}：既不是该用户，也没有 sudo/runuser" >&2
  exit 1
fi

# 以 ubuntu 身份执行；user 级 systemd 需要它自己的 runtime dir / dbus
as_app() {
  "${AS_APP[@]}" \
    HOME="${APP_HOME}" \
    XDG_RUNTIME_DIR="/run/user/${APP_UID}" \
    DBUS_SESSION_BUS_ADDRESS="unix:path=/run/user/${APP_UID}/bus" \
    "$@"
}

echo "--> ssh user=$(id -un) app user=${APP_USER} (uid ${APP_UID}, home ${APP_HOME})"

"${SUDO[@]}" mkdir -p "${REMOTE_DIR}/accounts" /var/www/certbot
"${SUDO[@]}" chown -R "${APP_USER}:${APP_USER}" "${REMOTE_DIR}"

# user 服务要在无登录会话时常驻，且 /run/user/<uid> 必须先存在
if command -v loginctl >/dev/null 2>&1; then
  "${SUDO[@]}" loginctl enable-linger "${APP_USER}" 2>/dev/null || true
fi
for _ in 1 2 3 4 5 6 7 8 9 10; do
  [[ -d "/run/user/${APP_UID}" ]] && break
  sleep 1
done

echo "==> venv + deps (as ${APP_USER})"
as_app bash -lc "cd '${REMOTE_DIR}' && { [[ -d .venv ]] || python3 -m venv .venv; } && .venv/bin/pip install -U pip -q && .venv/bin/pip install -r requirements.txt -q"

# .env：只补齐缺的键，不覆盖运维手工填过的 captcha key / 代理 / 恢复邮箱配置。
# admin key 留空且远端也没有时现场生成一枚，免得 Mailbox API 上线即无人可用。
echo "==> 渲染 ${ENV_FILE}（仅补齐缺失项）"
as_app "ADMIN_KEY=${ADMIN_KEY}" python3 - "${ENV_FILE}" <<'PY'
import os
import secrets
import sys
from pathlib import Path

path = Path(sys.argv[1])
lines = path.read_text(encoding="utf-8").splitlines() if path.exists() else []


def current(key: str) -> str:
    for line in lines:
        stripped = line.strip()
        if stripped.startswith(f"{key}="):
            return stripped.split("=", 1)[1].strip()
    return ""


def upsert(key: str, value: str) -> None:
    for i, line in enumerate(lines):
        if line.strip().startswith(f"{key}="):
            lines[i] = f"{key}={value}"
            return
    lines.append(f"{key}={value}")


admin_key = (os.environ.get("ADMIN_KEY") or "").strip()
source = "来自 .deploy/local.env"
if not admin_key:
    admin_key = current("OUTLOOK_MAILBOX_API_ADMIN_KEY")
    source = "沿用远端已有"
if not admin_key:
    admin_key = "mbxadmin_" + secrets.token_urlsafe(32)
    source = "本次新生成"

if not lines:
    lines = ["# 由 scripts/deploy_to_38.sh 渲染。手工加的键会被保留，不要提交进 git。"]

upsert("OUTLOOK_MAILBOX_API_ENABLED", "1")
upsert("OUTLOOK_MAILBOX_API_ADMIN_KEY", admin_key)
if not current("OUTLOOK_MAILBOX_API_SESSION_HOURS"):
    upsert("OUTLOOK_MAILBOX_API_SESSION_HOURS", "24")

path.write_text("\n".join(lines) + "\n", encoding="utf-8")
path.chmod(0o600)
print(f"admin key: {source}")
PY
"${SUDO[@]}" chown "${APP_USER}:${APP_USER}" "${ENV_FILE}"
"${SUDO[@]}" chmod 600 "${ENV_FILE}"

# 预检：宁可在这里失败，也不要让 systemd 反复拉起一个 import 就崩的进程
as_app bash -lc "cd '${REMOTE_DIR}' && .venv/bin/python -c 'from webapp.server import app; print(\"app import ok\")'"

echo "==> systemd (user service under ${APP_USER})"
UNIT_DIR="${APP_HOME}/.config/systemd/user"
"${SUDO[@]}" mkdir -p "${UNIT_DIR}"
# 口令可能含 sed 元字符，用 python 做字面替换
CONSOLE_PASSWORD="${CONSOLE_PASSWORD}" python3 - \
  "${REMOTE_DIR}/deploy/outlook-console.user.service" \
  "${UNIT_DIR}/outlook-console.service" <<'PY'
import os
import sys
from pathlib import Path

src, dst = Path(sys.argv[1]), Path(sys.argv[2])
text = src.read_text(encoding="utf-8").replace(
    "__OUTLOOK_CONSOLE_PASSWORD__", os.environ.get("CONSOLE_PASSWORD", "")
)
dst.write_text(text, encoding="utf-8")
dst.chmod(0o600)
PY
"${SUDO[@]}" chown -R "${APP_USER}:${APP_USER}" "${APP_HOME}/.config"

# 关掉历史上可能存在的 system 级同名服务，避免两个实例抢 8890
"${SUDO[@]}" systemctl disable --now outlook-console 2>/dev/null || true
"${SUDO[@]}" rm -f /etc/systemd/system/outlook-console.service 2>/dev/null || true
"${SUDO[@]}" systemctl daemon-reload 2>/dev/null || true

as_app systemctl --user daemon-reload
as_app systemctl --user enable outlook-console
as_app systemctl --user restart outlook-console

if [[ "${SKIP_NGINX}" != "1" ]] && command -v nginx >/dev/null 2>&1; then
  CERT="/etc/letsencrypt/live/${DOMAIN}/fullchain.pem"
  AVAIL="/etc/nginx/sites-available/${DOMAIN}"
  if [[ ! -f "${CERT}" ]]; then
    # 正式 vhost 里写着还不存在的证书路径，nginx -t 过不了，certbot 的 webroot
    # 校验也就落不到 /var/www/certbot。所以先上一份只监听 80 的临时 vhost。
    echo "==> nginx（临时 HTTP vhost，供 certbot 校验）"
    "${SUDO[@]}" tee "${AVAIL}" >/dev/null <<NGINX_BOOTSTRAP
server {
    listen 80;
    listen [::]:80;
    server_name ${DOMAIN};
    location /.well-known/acme-challenge/ {
        root /var/www/certbot;
    }
    location / {
        proxy_pass http://127.0.0.1:${PORT};
        proxy_http_version 1.1;
        proxy_set_header Host \$host;
        proxy_set_header Authorization \$http_authorization;
        proxy_pass_header Authorization;
    }
}
NGINX_BOOTSTRAP
    "${SUDO[@]}" ln -sf "${AVAIL}" "/etc/nginx/sites-enabled/${DOMAIN}"
    "${SUDO[@]}" nginx -t && "${SUDO[@]}" systemctl reload nginx
    echo "==> 签发 SSL 证书"
    "${SUDO[@]}" certbot certonly --webroot -w /var/www/certbot -d "${DOMAIN}" \
      --non-interactive --agree-tos -m admin@coolhs.com || true
  fi

  if [[ -f "${CERT}" ]]; then
    echo "==> nginx（正式 HTTPS vhost）"
    "${SUDO[@]}" cp "${REMOTE_DIR}/deploy/nginx-outlook.conf" "${AVAIL}"
    if [[ "${DOMAIN}" != "outlook.coolhs.com" ]]; then
      "${SUDO[@]}" sed -i "s/outlook\.coolhs\.com/${DOMAIN}/g" "${AVAIL}"
    fi
    "${SUDO[@]}" ln -sf "${AVAIL}" "/etc/nginx/sites-enabled/${DOMAIN}"
    "${SUDO[@]}" nginx -t && "${SUDO[@]}" systemctl reload nginx
  else
    echo "证书没签下来，暂时留着 HTTP vhost（先确认 ${DOMAIN} 的 DNS 指到本机再重跑）" >&2
  fi
else
  echo "==> 跳过 nginx（未安装 nginx 或 OUTLOOK_SKIP_NGINX=1）"
fi

echo "==> 本地探活"
sleep 3
as_app systemctl --user is-active outlook-console
curl -fsS "http://127.0.0.1:${PORT}/api/health" && echo ""
# 运维台的门：不带口令应当 302 到登录页
echo "GET / (未登录) -> $(curl -s -o /dev/null -w '%{http_code}' "http://127.0.0.1:${PORT}/")"
# Mailbox API：admin key 应当直接通。key 经 0600 的头文件交给 curl（-H @file），
# 不进 argv，免得在 ps 里一闪而过。
HDR="$(mktemp)"
chmod 600 "${HDR}"
printf 'Authorization: Bearer %s\n' \
  "$("${SUDO[@]}" grep -m1 '^OUTLOOK_MAILBOX_API_ADMIN_KEY=' "${ENV_FILE}" | cut -d= -f2-)" > "${HDR}"
echo "GET /api/v1/health (admin key) -> $(curl -s -o /dev/null -w '%{http_code}' -H "@${HDR}" "http://127.0.0.1:${PORT}/api/v1/health")"
rm -f "${HDR}"
REMOTE_SCRIPT
} | "${SSH[@]}" "${REMOTE}" bash -s

echo ""
echo "==> 域名探活"
curl -s -o /dev/null -w "https://${DOMAIN}/api/health -> %{http_code}\n" "https://${DOMAIN}/api/health" || true
curl -s -o /dev/null -w "https://${DOMAIN}/login.html -> %{http_code}\n" "https://${DOMAIN}/login.html" || true

echo ""
echo "部署完成。"
echo "  运维台: https://${DOMAIN}/  （口令见 .deploy/local.env 的 OUTLOOK_CONSOLE_PASSWORD）"
echo "  Mailbox API admin key 取法（不回显在这里）:"
echo "    ssh ${REMOTE} \"grep '^OUTLOOK_MAILBOX_API_ADMIN_KEY=' ${REMOTE_DIR}/.env\""
