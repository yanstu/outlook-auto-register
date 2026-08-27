"""用 Graph / Outlook REST 四段式令牌读信（绕开 IMAP 协议开关，新号注册完即用）。

结论依据（见 check_imap / 交付说明）：
- Graph 令牌（Mail.ReadWrite / Mail.Send / User.Read）→ GET graph.microsoft.com/v1.0/me/messages
- Outlook REST 令牌（outlook.office.com/Mail.Read 等）→ GET outlook.office.com/api/v2.0/me/messages
两者都不依赖邮箱 IMAP/POP 开关，新号立即可读；IMAP 令牌在新号上会
「User is authenticated but not connected」（协议未开启）。
"""
from __future__ import annotations

import html as _html
import logging
import re
import time
from typing import Optional

import requests
from requests.exceptions import ConnectionError, ProxyError, SSLError, Timeout

from .constants import (
    GRAPH_MAIL_SCOPE,
    LOGIN_CLIENT_ID,
    LOGIN_SCOPE,
    MAIL_CLIENT_ID,
    MAIL_REDIRECT_URI,
    OUTLOOK_REST_SCOPE,
)

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
OUTLOOK_REST_BASE = "https://outlook.office.com/api/v2.0"

_SECURITY_SENDERS = (
    "accountprotection.microsoft.com",
    "account-security-noreply",
    "microsoftonline.com",
    "microsoft.com",
)

_TRANSIENT_EXC = (ConnectionError, SSLError, Timeout, ProxyError)


def is_transient_error(exc: BaseException) -> bool:
    """网络/SSL/代理抖动：不应把账号判为失活。"""
    if isinstance(exc, _TRANSIENT_EXC):
        return True
    msg = str(exc).lower()
    return any(
        k in msg
        for k in (
            "ssl", "connection", "timeout", "timed out", "max retries",
            "connection reset", "connection aborted", "eof occurred",
        )
    )


def _http_get(url: str, *, headers: dict, proxies: Optional[dict], timeout: int = 30) -> requests.Response:
    try:
        return requests.get(url, headers=headers, proxies=proxies, timeout=timeout)
    except _TRANSIENT_EXC as exc:
        raise RuntimeError(f"transient_network:{exc}") from exc


def refresh_token_for(
    refresh_token: str,
    scope: str = "",
    *,
    client_id: str = MAIL_CLIENT_ID,
    proxy_url: str = "",
    retries: int = 3,
) -> dict:
    """refresh_token → token 响应。scope 为空则不带（返回令牌被授予的原始 scope，最稳）。"""
    proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None
    data = {
        "client_id": client_id,
        "refresh_token": refresh_token,
        "redirect_uri": MAIL_REDIRECT_URI,
        "grant_type": "refresh_token",
    }
    if scope:
        data["scope"] = scope
    last_exc: Optional[BaseException] = None
    for attempt in range(max(1, retries)):
        try:
            r = requests.post(
                "https://login.microsoftonline.com/common/oauth2/v2.0/token",
                data=data,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                proxies=proxies,
                timeout=30,
            )
            try:
                return r.json()
            except ValueError:
                return {"error": "non_json", "status": r.status_code, "body": r.text[:300]}
        except _TRANSIENT_EXC as exc:
            last_exc = exc
            if attempt + 1 < retries:
                time.sleep(1.0 * (attempt + 1))
                continue
    desc = str(last_exc or "network error")[:200]
    logger.warning("读信令牌网络异常（已重试 %d 次）: %s", retries, desc)
    return {
        "error": "transient_network",
        "error_description": desc,
        "transient": True,
    }


def token_alive(refresh_token: str, *, proxy_url: str = "") -> bool:
    """快速校验老号 refresh_token 是否仍可换 access_token（避免向读不了的老号发码）。"""
    data = refresh_token_for(refresh_token, "", proxy_url=proxy_url)
    return bool(data.get("access_token"))


def verify_mail_readable(
    refresh_token: str,
    *,
    client_id: str = MAIL_CLIENT_ID,
    proxy_url: str = "",
) -> dict:
    """注册产出前自检：这枚 refresh_token 能否真正读信。

    按令牌实际 scope 路由——Graph 令牌打 graph.microsoft.com，IMAP/REST 令牌打
    outlook.office.com/api/v2.0。只授 IMAP/POP/SMTP 的新号会 401（IMAP 协议默认关、
    REST 缺 Mail.Read）→ readable=False，用于产出前拦掉“建成但读不了”的废号。

    返回 {readable, resource, status, display_name, scope, reason?, detail?}。
    """
    data = refresh_token_for(refresh_token, "", client_id=client_id, proxy_url=proxy_url)
    at = data.get("access_token", "")
    if not at:
        return {
            "readable": False,
            "reason": "refresh_failed",
            "detail": str(data.get("error_description") or data.get("error") or "")[:120],
        }
    scope = data.get("scope", "")
    resource = _resource_of(scope)
    if resource == "outlook":
        url = OUTLOOK_REST_BASE + "/me"
    else:  # graph 或 unknown 一律按 Graph 试
        url = GRAPH_BASE + "/me?$select=userPrincipalName,displayName"
    proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None
    try:
        r = requests.get(url, headers={"Authorization": "Bearer " + at}, proxies=proxies, timeout=30)
    except Exception as exc:  # noqa: BLE001
        return {"readable": False, "reason": "network", "detail": str(exc)[:120], "resource": resource}
    display_name = ""
    if r.status_code == 200:
        try:
            j = r.json()
            display_name = j.get("DisplayName") or j.get("displayName") or ""
        except ValueError:
            pass
    return {
        "readable": r.status_code == 200,
        "resource": resource,
        "status": r.status_code,
        "display_name": display_name,
        "scope": scope[:200],
    }


def _resource_of(scope: str) -> str:
    """按返回的 scope 判定令牌资源：graph / outlook / unknown。"""
    s = (scope or "").lower()
    if "graph.microsoft.com" in s:
        return "graph"
    if "outlook.office.com" in s:
        return "outlook"
    # Graph 令牌常以短名返回（Mail.ReadWrite Mail.Send User.ReadWrite），无 host
    if "mail.readwrite" in s or "mail.read" in s or "user.read" in s:
        return "graph"
    return "unknown"


def _access_token(refresh_token: str, mode: str, proxy_url: str) -> str:
    """按 mode 请求对应 scope；失败则回退无 scope 刷新（拿原始授予 scope 的令牌）。"""
    scope = GRAPH_MAIL_SCOPE if mode == "graph" else OUTLOOK_REST_SCOPE
    data = refresh_token_for(refresh_token, scope, proxy_url=proxy_url)
    tok = data.get("access_token", "")
    if tok:
        return tok
    # 回退：无 scope 刷新，仅当资源匹配才用
    data = refresh_token_for(refresh_token, "", proxy_url=proxy_url)
    tok = data.get("access_token", "")
    if tok and _resource_of(data.get("scope", "")) == (mode if mode == "graph" else "outlook"):
        return tok
    if not tok:
        logger.error(
            "刷新 %s access_token 失败: %s %s",
            mode, data.get("error"), str(data.get("error_description", ""))[:120],
        )
    return tok if _resource_of(data.get("scope", "")) == ("graph" if mode == "graph" else "outlook") else ""


def list_messages(
    refresh_token: str,
    *,
    mode: str = "graph",
    top: int = 10,
    folder: str = "inbox",
    proxy_url: str = "",
) -> list[dict]:
    """读取最近邮件（graph 或 outlook_rest）。返回标准化 [{subject, from, received, body}]。"""
    at = _access_token(refresh_token, mode, proxy_url)
    if not at:
        return []
    proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None
    headers = {"Authorization": "Bearer " + at}
    if mode == "graph":
        url = (
            f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
            f"?$top={top}&$orderby=receivedDateTime desc"
            f"&$select=subject,from,receivedDateTime,bodyPreview,body"
        )
    else:
        url = (
            f"{OUTLOOK_REST_BASE}/me/MailFolders/{folder}/messages"
            f"?$top={top}&$orderby=ReceivedDateTime desc"
            f"&$select=Subject,From,ReceivedDateTime,BodyPreview,Body"
        )
    r = requests.get(url, headers=headers, proxies=proxies, timeout=30)
    if r.status_code != 200:
        logger.error("%s 读信失败 status=%s body=%s", mode, r.status_code, r.text[:200])
        return []
    out = []
    for m in r.json().get("value", []):
        if mode == "graph":
            frm = (((m.get("from") or {}).get("emailAddress") or {}).get("address")) or ""
            out.append({
                "subject": m.get("subject", ""),
                "from": frm,
                "received": m.get("receivedDateTime", ""),
                "preview": m.get("bodyPreview", ""),
                "body": (m.get("body") or {}).get("content") or m.get("bodyPreview", ""),
            })
        else:
            frm = (((m.get("From") or {}).get("EmailAddress") or {}).get("Address")) or ""
            out.append({
                "subject": m.get("Subject", ""),
                "from": frm,
                "received": m.get("ReceivedDateTime", ""),
                "preview": m.get("BodyPreview", ""),
                "body": (m.get("Body") or {}).get("Content") or m.get("BodyPreview", ""),
            })
    return out


def _clean_text(raw: str) -> str:
    """去掉 <style>/<script>/HTML 标签与 #hex 颜色，避免把模板灰色 #707070 当验证码。"""
    if not raw:
        return ""
    t = re.sub(r"<style[^>]*>.*?</style>", " ", raw, flags=re.I | re.S)
    t = re.sub(r"<script[^>]*>.*?</script>", " ", t, flags=re.I | re.S)
    t = re.sub(r"#[0-9a-fA-F]{3,8}\b", " ", t)          # hex 颜色 #707070 等
    t = re.sub(r"(?i)(?:color|background|fill)\s*[:=]\s*[^;\"'>]{0,20}", " ", t)
    t = re.sub(r"<[^>]+>", " ", t)                       # 其余标签
    t = _html.unescape(t)
    return re.sub(r"\s+", " ", t).strip()


def _pick_code(text: str) -> str:
    if not text:
        return ""
    # 优先「关键词 + 数字」（安全代码：123456 / your code is 123456）
    m = re.search(r"(?:code|代码|verification|安全代码|verify)\D{0,24}?(\d{4,8})", text, re.I)
    if m:
        return m.group(1)
    nums = re.findall(r"(?<!\d)(\d{4,8})(?!\d)", text)
    long_codes = [n for n in nums if len(n) >= 6]
    if long_codes:
        return long_codes[0]
    for n in nums:
        if not (len(n) == 4 and 1990 <= int(n) <= 2035):
            return n
    return nums[0] if nums else ""


def _extract_code(subject: str, body: str, preview: str = "") -> str:
    """从安全码邮件抽取验证码：subject → preview(纯文本) → 清洗后的 body。"""
    for text in (subject, preview, _clean_text(body)):
        code = _pick_code(text or "")
        if code:
            return code
    return ""


def read_security_code(
    refresh_token: str,
    *,
    mode: str = "graph",
    timeout: int = 150,
    proxy_url: str = "",
    since_iso: str = "",
) -> str:
    """轮询读取微软安全验证码（Graph/REST）。返回验证码或空串。"""
    subj_hints = ("security code", "verification code", "verify", "安全代码", "验证码",
                  "验证你的", "single-use code", "one-time")
    deadline = time.time() + timeout
    while time.time() < deadline:
        for folder in ("inbox", "junkemail"):
            msgs = list_messages(refresh_token, mode=mode, top=15, folder=folder, proxy_url=proxy_url)
            for m in msgs:
                sender_ok = any(s in (m["from"] or "").lower() for s in _SECURITY_SENDERS)
                subj_ok = any(h in (m["subject"] or "").lower() for h in subj_hints)
                if not (sender_ok or subj_ok):
                    continue
                if since_iso and m.get("received") and m["received"] < since_iso:
                    continue
                code = _extract_code(m["subject"], m["body"], m.get("preview", ""))
                if code:
                    logger.info(
                        "Graph/REST 读到安全验证码 %s (from=%s subject=%s received=%s preview=%s)",
                        code, m.get("from"), (m.get("subject") or "")[:40],
                        (m.get("received") or "")[:19], (m.get("preview") or "")[:60],
                    )
                    return code
        time.sleep(5)
    logger.error("Graph/REST 等待安全验证码超时(%ss)", timeout)
    return ""


def probe_token(email: str, refresh_token: str, *, proxy_url: str = "") -> dict:
    """自动探测四段式令牌可用方式：graph / outlook_rest / imap-only。

    返回 {mode_ok: [...], detail: {...}}，供 check_imap 汇报。
    """
    result: dict = {"email": email, "usable": [], "detail": {}}
    proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None

    # 1) 无 scope 刷新，拿令牌被授予的原始资源令牌，命中对应 API
    base = refresh_token_for(refresh_token, "", proxy_url=proxy_url)
    if base.get("transient"):
        result["transient"] = True
        result["detail"]["refresh"] = str(base.get("error_description", "network"))[:80]
        return result
    at0 = base.get("access_token", "")
    if not at0:
        result["detail"]["refresh"] = str(base.get("error_description", base.get("error")))[:80]
        return result
    res = _resource_of(base.get("scope", ""))
    result["detail"]["granted_scope"] = base.get("scope", "")
    try:
        if res == "graph":
            r = _http_get(
                GRAPH_BASE + "/me/messages?$top=1&$select=subject",
                headers={"Authorization": "Bearer " + at0}, proxies=proxies,
            )
            result["detail"]["graph"] = r.status_code
            if r.status_code == 200:
                result["usable"].append("graph")
        elif res == "outlook":
            r = _http_get(
                OUTLOOK_REST_BASE + "/me/messages?$top=1&$select=Subject",
                headers={"Authorization": "Bearer " + at0}, proxies=proxies,
            )
            result["detail"]["outlook_rest"] = r.status_code
            if r.status_code == 200:
                result["usable"].append("outlook_rest")
    except RuntimeError as exc:
        if str(exc).startswith("transient_network:"):
            result["transient"] = True
            result["detail"]["refresh"] = str(exc)[len("transient_network:"):][:80]
            return result
        raise

    # 2) 若默认资源不是 graph，再尝试显式换 graph（部分令牌被授予了多资源）
    if "graph" not in result["usable"]:
        scoped = refresh_token_for(refresh_token, GRAPH_MAIL_SCOPE, proxy_url=proxy_url)
        if scoped.get("transient"):
            result["transient"] = True
            result["detail"]["refresh"] = str(scoped.get("error_description", "network"))[:80]
            return result
        atg = scoped.get("access_token", "")
        if atg:
            try:
                r = _http_get(
                    GRAPH_BASE + "/me/messages?$top=1&$select=subject",
                    headers={"Authorization": "Bearer " + atg}, proxies=proxies,
                )
                result["detail"]["graph"] = r.status_code
                if r.status_code == 200:
                    result["usable"].append("graph")
            except RuntimeError as exc:
                if str(exc).startswith("transient_network:"):
                    result["transient"] = True
                    result["detail"]["refresh"] = str(exc)[len("transient_network:"):][:80]
                    return result
                raise
    return result


def probe_login_token(
    refresh_token: str,
    *,
    client_id: str = LOGIN_CLIENT_ID,
    scope: str = LOGIN_SCOPE,
    proxy_url: str = "",
) -> dict:
    """探测双令牌 token#2（登录授权）是否可用。

    判定标准：能用该 refresh_token 换到 access_token（可选带 id_token）→ 可继续做
    “用微软账号登录”的 SSO。返回 {usable, id_token, granted_scope, error}。
    """
    out: dict = {"usable": False, "id_token": False, "granted_scope": "", "error": ""}
    data = refresh_token_for(refresh_token, scope, client_id=client_id, proxy_url=proxy_url)
    at = data.get("access_token", "")
    if not at:
        # 回退：不带 scope 刷新（取原始授予 scope）
        data = refresh_token_for(refresh_token, "", client_id=client_id, proxy_url=proxy_url)
        at = data.get("access_token", "")
    if at:
        out["usable"] = True
        out["granted_scope"] = data.get("scope", "")
        out["id_token"] = bool(data.get("id_token"))
    else:
        out["error"] = str(data.get("error_description", data.get("error")))[:100]
    return out
