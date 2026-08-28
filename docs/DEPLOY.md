# 部署到 38 服务器（outlook.coolhs.com）

一条命令把整个仓库推上去、装依赖、起服务、配 nginx 和证书：

```bash
cd outlook-auto-register && ./scripts/deploy_to_38.sh
```

| 项 | 值 |
| --- | --- |
| 域名 | `https://outlook.coolhs.com` |
| 远端目录 | `/home/ubuntu/outlook-auto-register` |
| 进程 | `ubuntu` 用户的 systemd user 服务 `outlook-console` |
| 监听 | `127.0.0.1:8890`（只走 nginx，不对外裸奔） |
| nginx | `/etc/nginx/sites-available/outlook.coolhs.com` |
| 证书 | Let's Encrypt，certbot webroot（`/var/www/certbot`） |

## 一、准备凭据（只做一次）

```bash
cp .deploy/local.env.example .deploy/local.env
chmod 600 .deploy/local.env
$EDITOR .deploy/local.env
```

| 变量 | 说明 |
| --- | --- |
| `OUTLOOK_DEPLOY_HOST` | SSH 目标，形如 `root@38.76.160.89` |
| `OUTLOOK_DEPLOY_PASSWORD` | SSH 密码；填了走 sshpass，留空走 `~/.ssh` 密钥 |
| `OUTLOOK_DOMAIN` | 域名，默认 `outlook.coolhs.com` |
| `OUTLOOK_CONSOLE_PASSWORD` | 运维台登录口令，**不能为空** |
| `OUTLOOK_MAILBOX_API_ADMIN_KEY` | Mailbox API 的 admin key；留空则沿用远端已有的，远端也没有就现场生成 |
| `OUTLOOK_SKIP_NGINX` | 设 `1` 时只更新应用，不动 nginx 与证书 |

`local.env` 已在 `.gitignore` 里。真口令只允许留在这个文件和远端的 systemd 单元里，
**不要**写进脚本、文档、提交信息或对话记录。

## 二、脚本做了什么

1. **rsync** 源码，排除 `.venv/`、`accounts/`、`.git/`、`__pycache__/`、`.env`、`.deploy/`。
   `accounts/`（线上账号库与导出文件）和 `.env`（线上 captcha key / 代理 / admin key）
   都由服务器自己管，`--delete` 碰不到它们。
2. **venv + 依赖**：以 `ubuntu` 身份建 `.venv` 并装 `requirements.txt`。
   SSH 用 root 登录时脚本自动降权，不会在 `/home/ubuntu` 下留 root 属主的文件。
3. **渲染 `.env`**：只补齐缺的键（`OUTLOOK_MAILBOX_API_ENABLED`、
   `OUTLOOK_MAILBOX_API_ADMIN_KEY`、`OUTLOOK_MAILBOX_API_SESSION_HOURS`），
   运维手工填过的 captcha / 代理 / 恢复邮箱配置原样保留。文件权限 0600。
4. **预检**：远端 `import webapp.server`，import 就崩的话当场失败，
   而不是让 systemd 反复拉起一个起不来的进程。
5. **systemd user 服务** `outlook-console`（`deploy/outlook-console.user.service`），
   顺带 `enable-linger`，没人登录也常驻。
6. **nginx + certbot**：证书还没签过时先上一份只监听 80 的临时 vhost 让 webroot
   校验能过，签下来后再换成正式的 HTTPS vhost（`deploy/nginx-outlook.conf`）。
7. **探活**：远端 `/api/health`、未登录访问 `/` 应当 302、带 admin key 打
   `/api/v1/health` 应当 200；最后本地再打一次域名。

## 三、两道门，各管各的

| 入口 | 谁用 | 凭证 |
| --- | --- | --- |
| `/`、`/api/*` | 运维（浏览器 / curl） | `/login.html` 登录后的 httponly cookie，或 `X-Console-Password` 头 |
| `/api/v1/*` | 其他项目 | `Authorization: Bearer mbx_sk_…`（见 [MAILBOX_API.md](MAILBOX_API.md)） |
| `/api/health` | 探活 | 无 |

`OUTLOOK_CONSOLE_PASSWORD` 非空时，`webapp/server.py` 的 middleware 会在路由之前
拦下所有请求：浏览器打开页面会 302 到 `/login.html`，`/api/*` 直接 401 JSON。

**Mailbox API 不受运维口令影响**。它给别的项目调用，再叠一层运维口令会让所有调用方
一起 401，所以 `/api/v1` 整段绕开这道门，只认自己的 Bearer key。反过来也一样：
运维台的会话 cookie 打不开 `/api/v1`。

```bash
# 运维台
curl -s -H 'X-Console-Password: <口令>' https://outlook.coolhs.com/api/config

# Mailbox API
curl -s -H "Authorization: Bearer $KEY" https://outlook.coolhs.com/api/v1/health
```

## 四、取 Mailbox API 的 admin key

key 只存在于远端 `.env`（0600，仅 `ubuntu` 可读），本地脚本和日志都不回显：

```bash
ssh root@38.76.160.89 "grep '^OUTLOOK_MAILBOX_API_ADMIN_KEY=' /home/ubuntu/outlook-auto-register/.env"
```

拿到 admin key 后，给每个调用方单独发一枚窄权限的 key，不要把 admin key 直接给出去：

```bash
ssh root@38.76.160.89 "cd /home/ubuntu/outlook-auto-register && sudo -u ubuntu .venv/bin/python scripts/mailbox_api_key.py create \
  --name kimi-register --scopes mailboxes:read,fields:basic,messages:read,otp:read"
```

换 admin key：改 `.deploy/local.env` 里的 `OUTLOOK_MAILBOX_API_ADMIN_KEY` 再跑一次部署。
旧 key 不会自动失效，要停用得 `scripts/mailbox_api_key.py revoke <id>`。

## 五、日常运维

```bash
SSH="ssh root@38.76.160.89"
AS="sudo -u ubuntu XDG_RUNTIME_DIR=/run/user/$($SSH id -u ubuntu)"

$SSH "$AS systemctl --user status outlook-console"
$SSH "$AS systemctl --user restart outlook-console"
$SSH "$AS journalctl --user -u outlook-console -n 200 --no-pager"
```

只改了代码、不想动 nginx：

```bash
OUTLOOK_SKIP_NGINX=1 ./scripts/deploy_to_38.sh
```

## 六、排查

**服务起不来** — 先看 `journalctl --user -u outlook-console`。多半是依赖没装全或
`.env` 写坏了；远端手动跑一次 `cd /home/ubuntu/outlook-auto-register &&
.venv/bin/python -c 'from webapp.server import app'` 能直接看到 traceback。

**证书没签下来** — 脚本会保留临时的 HTTP vhost 并打印提示。确认 `outlook.coolhs.com`
的 DNS 指到这台机器（Cloudflare 代理开着也行，HTTP-01 校验会穿透到源站的
`/.well-known/acme-challenge/`），然后重跑部署。手工签：

```bash
ssh root@38.76.160.89 "certbot certonly --webroot -w /var/www/certbot -d outlook.coolhs.com"
```

**域名 502/503，但远端 `curl 127.0.0.1:8890/api/health` 正常** — nginx 没 reload，
或 `sites-enabled` 的软链没建上。`nginx -t && systemctl reload nginx` 看报什么。

**其他项目调 `/api/v1` 全是 401** — 先确认 nginx 没吞掉 `Authorization`
（`deploy/nginx-outlook.conf` 里有 `proxy_set_header Authorization` +
`proxy_pass_header Authorization`），再用 `GET /api/v1/auth/me` 自检那枚 key。

**运维台反复跳登录页** — 口令改过之后旧会话 cookie 立即失效（cookie 就是口令的 HMAC），
重新登录即可。

**8890 被占** — systemd 单元里有 `ExecStartPre=fuser -k 8890/tcp`，正常情况下会自己让位；
仍占着就看是不是有人手工起了一个 uvicorn。
