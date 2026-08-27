#!/usr/bin/env python3
"""ss_register / bit_register 到达 proofs/Add 之后的续跑桥接。

默认：恢复邮箱在浏览器里填（#EmailAddress → 收码 → #iOttText），本模块只交接 cookie
跑 Thunderbird OAuth。BIT_BROWSER_PROOFS=0 时回退为 requests AddProof。
"""
from __future__ import annotations

import json
import os
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

# 让 `outlook_api_reg` 包可被 import（ss_post.py 在 px_solver/ 下，包在项目根）。
_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_HERE)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

# 真实 Windows Chrome UA（与建号浏览器一致，交接 cookie 时保持指纹连续）。
_BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
               "(KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36")


def _load_env() -> None:
    """加载项目根 .env（proofs 恢复邮箱池 / IMAP host / 代理等约定都在其中）。"""
    try:
        from dotenv import load_dotenv
        load_dotenv(os.path.join(_PROJECT_ROOT, ".env"))
    except Exception:  # noqa: BLE001
        pass


def _resolve_scope(token_mode: str) -> tuple[str, str]:
    """按产出模式选 OAuth scope；默认 login_exe = Thunderbird IMAP/POP/SMTP（HAR 坐实）。

    返回 (mode, scope)。可用 OUTLOOK_MAIL_TOKEN_MODE 覆盖（graph/outlook_rest/imap/login_exe/recovery）。
    """
    from outlook_api_reg.constants import (
        IMAP_MAIL_SCOPE,
        MAIL_SCOPE_BY_MODE,
        normalize_token_mode,
    )

    raw = (token_mode or os.environ.get("OUTLOOK_MAIL_TOKEN_MODE", "") or "login_exe")
    mode = normalize_token_mode(raw)
    scope = MAIL_SCOPE_BY_MODE.get(mode, IMAP_MAIL_SCOPE)
    return mode, scope


def _seed_cookies(session, cookies: list[dict]) -> int:
    """把浏览器 ctx.cookies() 灌进 requests 会话的 cookiejar。

    cookies 元素形如 Playwright 的 {name,value,domain,path,...}。__Host-* 是 host-only cookie，
    domain 无前导点，requests 精确匹配该 host；.live.com 之类带点 cookie 走后缀匹配。
    """
    import requests

    n = 0
    for c in cookies or []:
        name = c.get("name")
        value = c.get("value")
        if not name:
            continue
        domain = (c.get("domain") or "").strip()
        path = c.get("path") or "/"
        try:
            if domain:
                ck = requests.cookies.create_cookie(
                    name=name, value=value or "", domain=domain, path=path,
                )
                session.cookies.set_cookie(ck)
            else:
                session.cookies.set(name, value or "", path=path)
            n += 1
        except Exception:  # noqa: BLE001
            # 单个 cookie 失败不致命，跳过继续（宁可少一个也不要中断交接）
            continue
    return n


def _build_ctx(email: str, proofs_url: str, country: str):
    """构造最小 SignupSession：fetch_mail_oauth_code 只用到 uaid/mkt，其余给空串占位即可。"""
    from outlook_api_reg.constants import locale_for_country
    from outlook_api_reg.models import SignupSession

    mkt, lc = locale_for_country(country)
    return SignupSession(
        uaid=uuid.uuid4().hex,
        signup_url="https://signup.live.com/signup?lic=1",
        signup_page_url=proofs_url or "https://signup.live.com/signup?lic=1",
        cobrandid="", contextid="", opid="", bk="", sru="", canary="",
        mkt=mkt, lc=lc,
    )


class _ShimResp:
    """satisfy_proofs_with_external 只读 resp.text / resp.url；当 requests 重取 proofs 页
    未拿到 frmAddProof 时，用浏览器已渲染的 proofs HTML 兜底喂入（canary 走会话绑定）。"""

    def __init__(self, text: str, url: str):
        self.text = text or ""
        self.url = url or ""
        self.status_code = 200


def _export_combo(
    *, email: str, password: str, client_id: str, refresh_token: str,
    recovery_email: str, recovery_password: str, out_dir: str, extra: dict,
    log: Callable[..., None],
) -> dict:
    """六段活号导出：accounts/accounts_recovery.txt + 单号 JSON 快照。复用 RegisterResult.to_combo。"""
    from outlook_api_reg.models import RegisterResult

    result = RegisterResult(
        success=True, email=email, password=password,
        client_id=client_id, refresh_token=refresh_token,
        recovery_email=recovery_email, recovery_password=recovery_password,
        extra=extra,
    )
    combo6 = result.to_combo(recovery=True)

    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = email.replace("@", "_at_").replace(".", "_")
    snap = out / f"ss_{safe}_{ts}.json"
    snap.write_text(json.dumps({
        "email": email, "password": password, "client_id": client_id,
        "refresh_token": refresh_token, "recovery_email": recovery_email,
        "recovery_password": recovery_password, "combo_recovery": combo6,
        "source": "ss_register", "created_at": datetime.now().isoformat(),
        **extra,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    rec_path = out / "accounts_recovery.txt"
    with rec_path.open("a", encoding="utf-8") as fp:
        fp.write(combo6 + "\n")
    log("六段活号已导出 → %s（并追加 %s）", snap, rec_path)
    try:
        from outlook_api_reg.account_store import save_register_result
        save_register_result(result, batch_label="bit_register")
        log("账号已写入 SQLite accounts 表：%s", email)
    except Exception as exc:  # noqa: BLE001
        log("⚠️ SQLite 入库失败（JSON/txt 已落盘）: %s", exc)
    return {"combo_recovery": combo6, "combo_path": str(rec_path), "snapshot": str(snap)}


def finish_after_proofs(
    *,
    email: str,
    password: str,
    proofs_url: str,
    proofs_html: str,
    cookies: list[dict],
    proxy: Optional[str],
    country: str = "US",
    token_mode: str = "",
    out_dir: str = "accounts",
    log: Optional[Callable[..., None]] = None,
    proofs_done: bool = False,
    recovery_email: str = "",
    recovery_password: str = "",
    proofs_method: str = "",
) -> dict:
    """cookie 交接后跑 Thunderbird OAuth；默认还可在 HTTP 里绑 proofs。

    proofs_done=True 时跳过 AddProof（恢复邮箱已在浏览器里绑完），只换 token。
    """
    _load_env()
    if log is None:
        def log(msg, *a):  # type: ignore
            print("[ss_post] " + (msg % a if a else msg))

    info: dict = {
        "status": "error", "email": email, "recovery_email": "",
        "recovery_password": "", "refresh_token": "", "proofs_method": "",
        "final_url": "", "authorize_url": "",
    }

    try:
        from outlook_api_reg import external_recovery_pool as ext_pool
        from outlook_api_reg.constants import MAIL_CLIENT_ID
        from outlook_api_reg.http_session import OutlookHttpSession
        from outlook_api_reg.post_register import (
            exchange_code_for_token,
            fetch_mail_oauth_code,
            satisfy_proofs_with_external,
        )

        mode, scope = _resolve_scope(token_mode)
        info["token_mode"] = mode
        info["scope"] = scope
        log("续跑 proofs+OAuth：token_mode=%s client_id=%s country=%s proofs_done=%s",
            mode, MAIL_CLIENT_ID, country, proofs_done)

        if not proofs_done and not ext_pool.external_pool_enabled():
            backend = ext_pool.recovery_backend()
            info["status"] = "proofs_failed"
            if backend == "cf_domain":
                info["note"] = (
                    "CF 域名恢复邮箱后端未配置：需 OUTLOOK_CF_WORKER_API_URL + "
                    "OUTLOOK_CF_WORKER_ADMIN_TOKEN + OUTLOOK_CF_DOMAIN。proofs 无法绑定恢复邮箱。"
                )
            elif backend == "coolhs_mail":
                info["note"] = (
                    "coolhs-mail 恢复邮箱后端未配置：需 COOLHS_MAIL_BASE_URL + "
                    "COOLHS_MAIL_API_TOKEN + COOLHS_MAIL_DOMAIN。proofs 无法绑定恢复邮箱。"
                )
            else:
                info["note"] = (
                    "外部恢复邮箱池未配置：需 OUTLOOK_EXTERNAL_RECOVERY_POOL_FILE(每行 email----password) "
                    "+ OUTLOOK_RECOVERY_IMAP_HOST。proofs 无法绑定恢复邮箱。"
                    "（或设 OUTLOOK_RECOVERY_BACKEND=coolhs_mail / cf_domain）"
                )
            log("⚠️ %s", info["note"])
            return info

        # 1) 交接：requests 会话 + 浏览器 UA + 浏览器 cookie
        http = OutlookHttpSession(proxy=proxy)
        http.session.headers["User-Agent"] = _BROWSER_UA
        seeded = _seed_cookies(http.session, cookies)
        log("已交接 %d 个浏览器 cookie 到 requests 会话（代理=%s）",
            seeded, (str(proxy)[:40] if proxy else "DIRECT"))

        if proofs_done:
            if not recovery_email:
                info["status"] = "proofs_failed"
                info["note"] = "浏览器声称已绑 proofs 但未带回 recovery_email"
                log("⚠️ %s", info["note"])
                return info
            info.update({
                "recovery_email": recovery_email,
                "recovery_password": recovery_password,
                "proofs_method": proofs_method or "browser_proofs",
                "final_url": (proofs_url or "")[:200],
            })
            log("浏览器已绑定恢复邮箱 %s，跳过 HTTP AddProof，只跑 OAuth", recovery_email)
        else:
            # 2) 重取 proofs 页，验证会话续上；取不到表单则用浏览器 HTML 兜底
            resp = None
            try:
                resp = http.get(proofs_url, allow_redirects=True)
                has_form = "frmAddProof" in (resp.text or "")
                log("requests 重取 proofs 页 status=%s frmAddProof=%s url=%s",
                    resp.status_code, has_form, (resp.url or "")[:90])
            except Exception as exc:  # noqa: BLE001
                log("requests 重取 proofs 页异常：%s", exc)
                has_form = False

            if not (resp is not None and "frmAddProof" in (resp.text or "")):
                if proofs_html and "frmAddProof" in proofs_html:
                    log("重取未见 frmAddProof，回退用浏览器已渲染 proofs HTML")
                    resp = _ShimResp(proofs_html, proofs_url)
                else:
                    info["status"] = "handoff_failed"
                    info["note"] = "cookie 交接后 requests 无法加载 proofs 表单（会话可能未续上，需检查 cookie/代理一致性）"
                    log("⚠️ %s", info["note"])
                    return info

            # 3) 绑定外部恢复邮箱（AddProof → 收 OTT → VerifyProof → SLT）
            ext = satisfy_proofs_with_external(http, resp, country=country)
            if ext is None:
                info["status"] = "proofs_failed"
                info["note"] = "外部恢复邮箱池均未满足 proofs（令牌/收码失败或全部占用）"
                log("⚠️ %s", info["note"])
                return info
            r_next, meta = ext
            info.update({
                "recovery_email": meta.get("recovery_email", ""),
                "recovery_password": meta.get("recovery_password", ""),
                "proofs_method": meta.get("proofs_method", "external_recovery"),
                "final_url": (r_next.url or "")[:200],
            })
            log("✅ proofs 已用外部恢复邮箱绑定：%s → 续跳 %s",
                info["recovery_email"], info["final_url"][:90])

        # 4) Thunderbird OAuth：authorize → code
        ctx = _build_ctx(email, proofs_url, country)
        code_info = fetch_mail_oauth_code(http, ctx, email, client_id=MAIL_CLIENT_ID, scope=scope)
        info["authorize_url"] = (code_info.get("authorize_url", "") or "")[:200]
        code = code_info.get("code", "")
        if not code:
            info["status"] = "oauth_no_code"
            info["note"] = "proofs 已过但 authorize 未拿到 code（consent/interrupt 打断或会话未登录）"
            log("⚠️ %s authorize_url=%s", info["note"], info["authorize_url"][:110])
            return info

        # 5) code → refresh_token
        tok = exchange_code_for_token(http, code, client_id=MAIL_CLIENT_ID, scope=scope)
        rt = tok.get("refresh_token", "")
        info["scope_granted"] = tok.get("scope", "")
        if not rt:
            info["status"] = "oauth_no_token"
            info["note"] = "拿到 code 但未换到 refresh_token（scope 缺 offline_access 或 consent 未过）"
            log("⚠️ %s granted_scope=%s", info["note"], info.get("scope_granted", ""))
            return info
        info["refresh_token"] = rt
        log("✅ 已换到 IMAP/REST refresh_token（scope=%s）", info["scope_granted"][:80])

        # 产出前自检：确认令牌真能读信（只授 IMAP/POP/SMTP 的新号会 401 → 废号）。
        try:
            from outlook_api_reg.graph_mail import verify_mail_readable
            from outlook_api_reg.proxy_utils import parse_proxy

            _cfg = parse_proxy(proxy) if proxy else None
            chk = verify_mail_readable(
                rt, client_id=MAIL_CLIENT_ID, proxy_url=(_cfg.url if _cfg else ""),
            )
            info["mail_readable"] = chk.get("readable")
            info["mail_check"] = chk
            if chk.get("readable"):
                log("✅ 产出自检：令牌可读信 resource=%s /me=%s name=%s",
                    chk.get("resource"), chk.get("status"), chk.get("display_name"))
            else:
                log("⚠️ 产出自检：令牌读信失败（%s %s）——该号可能不可用",
                    chk.get("status") or chk.get("reason"), chk.get("detail", ""))
        except Exception as exc:  # noqa: BLE001
            log("⚠️ 产出自检异常（不阻断产出）：%s", exc)

        # 6) 六段活号导出
        exp = _export_combo(
            email=email, password=password, client_id=MAIL_CLIENT_ID, refresh_token=rt,
            recovery_email=info["recovery_email"], recovery_password=info["recovery_password"],
            out_dir=out_dir,
            extra={
                "proofs_method": info["proofs_method"],
                "source": info.get("proofs_method") or "ss_register",
                "token_mode": mode,
                "scope_granted": info.get("scope_granted", ""),
                "final_url": info["final_url"],
                "mail_readable": info.get("mail_readable"),
                "mail_check": info.get("mail_check", {}),
            },
            log=log,
        )
        info.update(exp)
        info["status"] = "ok"
        return info

    except Exception as exc:  # noqa: BLE001
        info["status"] = "error"
        info["note"] = f"续跑异常：{exc!r}"
        log("续跑异常：%r", exc)
        return info
