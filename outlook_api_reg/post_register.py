from __future__ import annotations

import html as html_lib
import logging
import os
import re
import time
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import requests

from .constants import (
    ACCOUNT_BASE,
    DUAL_TOKEN,
    LOGIN_CLIENT_ID,
    LOGIN_SCOPE,
    LOGIN_TOKEN_ATTEMPTS,
    MAIL_CLIENT_ID,
    MAIL_REDIRECT_URI,
    MAIL_SCOPE,
)
from .http_session import OutlookHttpSession
from .models import SignupSession

logger = logging.getLogger(__name__)


def _skip_proofs_allowed() -> bool:
    """默认禁止 cancel 跳过 proofs（批量 skip 是 abuse 主因之一）。显式 OUTLOOK_SKIP_PROOFS=1 才允许。"""
    return os.environ.get("OUTLOOK_SKIP_PROOFS", "0").strip() == "1"


_INPUT_RE = re.compile(r"<input\b[^>]*>", re.I)
_FORM_RE = re.compile(r"<form\b[^>]*>(.*?)</form>", re.I | re.S)
_FORM_BY_ID_RE = re.compile(
    r'<form\b[^>]*\bid=["\'](?P<id>[^"\']+)["\'][^>]*>(?P<inner>.*?)</form>', re.I | re.S
)


def _attr(tag: str, name: str) -> str:
    m = re.search(rf'{name}\s*=\s*"([^"]*)"', tag, re.I) or re.search(
        rf"{name}\s*=\s*'([^']*)'", tag, re.I
    )
    return html_lib.unescape(m.group(1)) if m else ""


def _decode_js_str(s: str) -> str:
    """还原 $Config 里的 \\u0026 / \\u002f 等 JS 转义。"""
    try:
        return s.encode("utf-8").decode("unicode_escape")
    except Exception:  # noqa: BLE001
        return (
            s.replace("\\u0026", "&").replace("\\u003a", ":").replace("\\u002f", "/")
        )


def _config_str(body: str, key: str) -> str:
    m = re.search(rf'"{key}"\s*:\s*"((?:[^"\\]|\\.)*)"', body)
    return _decode_js_str(m.group(1)) if m else ""


def _find_skip_url(body: str) -> str:
    """账号安全信息插页（proofs/Add, ManageProofsV2）的「跳过/稍后」链接。

    页面是 $Config 驱动 SPA，cancel viewDef 的 url 即 oauth 续跳地址，GET 它即可跳过。
    """
    m = re.search(r'"cancel"\s*:\s*\{\s*"url"\s*:\s*"((?:[^"\\]|\\.)*)"', body)
    if m:
        return _decode_js_str(m.group(1))
    return ""


def _is_consent_page(body: str, url: str) -> bool:
    """消费者账号 OAuth「同意授权」确认页（app-consent-fabric）。

    抓包证据：login.live.com/oauth20_authorize.srf → fmHF 自动 POST 到
    account.live.com/Consent/Update（rd/pprid/ipt/uaid/client_id/scope）→ 返回本页，
    页面 $Config 带 sClientId / sCanary / sRawInputScopes，需再 POST ucaction=Yes 才发码。
    """
    low = url.lower()
    return ("sRawInputScopes" in body) or ("/consent/" in low and "sCanary" in body)


def submit_consent(http: OutlookHttpSession, resp: requests.Response) -> Optional[requests.Response]:
    """在同意授权确认页 POST ucaction=Yes 完成授权（对应抓包 Consent/Update 第 2 次 POST）。"""
    body = resp.text or ""
    if not _is_consent_page(body, resp.url or ""):
        return None
    canary = _config_str(body, "sCanary")
    client_id = _config_str(body, "sClientId")
    scope = _config_str(body, "sRawInputScopes")
    if not (canary and client_id and scope):
        logger.warning(
            "识别到同意页但缺字段: canary=%s client_id=%s scope=%s",
            bool(canary), bool(client_id), bool(scope),
        )
        return None
    action = resp.url  # account.live.com/Consent/Update?...&ru=login.live.com/oauth20_authorize.srf...
    logger.info("提交授权同意")
    return http.post(
        action,
        data={
            "ucaction": "Yes",
            "client_id": client_id,
            "scope": scope,
            "cscope": "",
            "canary": canary,
        },
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": "https://account.live.com",
            "Referer": action,
        },
        allow_redirects=True,
    )


def _is_kmsi_page(body: str) -> bool:
    """「保持登录状态?（Stay signed in / KMSI）」插页。"""
    if "urlPost" not in body:
        return False
    return ("KmsiInterrupt" in body) or ('"fShowKmsi":true' in body.replace(" ", "")) or (
        "iRemainingDaysToSkipKmsi" in body and '"sErrorCode":""' in body.replace(" ", "")
    )


def submit_kmsi(http: OutlookHttpSession, resp: requests.Response) -> Optional[requests.Response]:
    """KMSI 插页选择「否」(LoginOptions=3) 续跳；选是/否都能继续，取否避免额外绑定。"""
    body = resp.text or ""
    if not _is_kmsi_page(body):
        return None
    url_post = _config_str(body, "urlPost")
    if not url_post:
        return None
    action = urllib.parse.urljoin(resp.url or "", url_post)
    ft = _config_str(body, "sFT")
    ft_name = _config_str(body, "sFTName") or "PPFT"
    canary = _config_str(body, "canary")
    fields: dict[str, str] = {"LoginOptions": "1", "type": "28"}
    if ft:
        fields[ft_name] = ft
    if canary:
        fields["canary"] = canary
    logger.info("KMSI 插页续跳 (LoginOptions=1)")
    return http.post(action, data=fields, allow_redirects=True)


def _parse_form_by_id(body: str, form_id: str) -> Optional[dict]:
    """按 form id 解析出 {action, method, fields{name:value}}（proofs Add/Verify 用）。"""
    for m in _FORM_BY_ID_RE.finditer(body):
        if m.group("id") != form_id:
            continue
        open_tag = m.group(0)[: m.group(0).find(">") + 1]
        action = html_lib.unescape(_attr(open_tag, "action"))
        method = (_attr(open_tag, "method") or "post").lower()
        fields: dict[str, str] = {}
        for inp in _INPUT_RE.findall(m.group("inner")):
            n = _attr(inp, "name")
            if not n:
                continue
            # 多个同名 input（如 proof=OTT||.. / proof=CSS）：保留第一个「有实义」的值
            val = _attr(inp, "value")
            if n in fields and (fields[n] or not val):
                continue
            fields[n] = val
        return {"action": action, "method": method, "fields": fields}
    return None


def _iso_now_minus(seconds: int = 90) -> str:
    """当前 UTC 时刻（提前 seconds 秒，容忍时钟/投递抖动）→ Graph receivedDateTime 可比 ISO。"""
    ts = time.time() - seconds
    return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _is_proofs_add_page(body: str, url: str) -> bool:
    return ("frmAddProof" in body) or ("proofs/Add" in url and "EmailAddress" in body)


def _extract_success_slt(body: str, cur_url: str) -> tuple[str, str, dict]:
    """VerifyProof 成功页会带 frmSubmitSLT + 非空 slt（HAR 坐实）→ 返回 (slt, action, fields)。

    失败页 slt 为空/不存在 → 返回 ("", "", {})。
    """
    form = _parse_form_by_id(body, "frmSubmitSLT")
    slt = ""
    if form:
        slt = form["fields"].get("slt", "")
    if not slt:
        m = re.search(r'name=["\']slt["\']\s+value=["\']([^"\']+)["\']', body) or re.search(
            r'"slt"\s*:\s*"([^"]{8,})"', body
        )
        if m:
            slt = _decode_js_str(m.group(1))
    if not slt:
        return "", "", {}
    if form:
        action = urllib.parse.urljoin(cur_url, form["action"])
        fields = dict(form["fields"])
        fields["slt"] = slt
    else:
        action = "https://login.live.com/ppsecure/post.srf"
        fields = {"slt": slt}
    return slt, action, fields


def satisfy_proofs_with_pool(
    http: OutlookHttpSession,
    resp: requests.Response,
    *,
    proxy: Optional[str] = None,
    max_accounts: int = 6,
    country: str = "US",
) -> Optional[tuple[requests.Response, dict[str, str]]]:
    """撞到 account.live.com/proofs/Add 时，用收码池老号当**恢复邮箱**满足 proofs。

    协议（抓包坐实，见 Outlook抓包.har）：
      1) AddProof: POST frmAddProof(action 含 ru=oauth) 字段
         iProofOptions=Email & EmailAddress=<老号> & canary & action=AddProof & Phone*空
         → 微软把 OTT 验证码发到老号收件箱，返回 proofs/Verify 页(frmVerifyProof)。
      2) 读 OTT: 用老号 refresh_token 走 Graph 读安全码（time-filter 避免旧码）。
      3) VerifyProof: POST frmVerifyProof(action 含 epid) 字段
         iProofOptions=OTT||老号||Email||0||t & iOttText=<code> & action=VerifyProof
         & canary(新) & GeneralVerify=0 → 续跳 SLT → oauth authorize。
    成功返回续跳后的 response；全部老号失败返回 None。
    """
    from . import graph_mail
    from .proof_pool import iter_accounts

    body = resp.text or ""
    add_form = _parse_form_by_id(body, "frmAddProof")
    if not add_form:
        logger.warning("proofs 页未解析到 frmAddProof 表单，无法用收码池满足")
        return None
    add_action = urllib.parse.urljoin(resp.url or "", add_form["action"])
    add_canary = add_form["fields"].get("canary", "")
    read_proxy = os.environ.get("OUTLOOK_PROOF_READ_PROXY", "")  # 老号读信默认直连

    for acct in iter_accounts(limit=max_accounts):
        logger.info("proofs 收码池：选用恢复老号 %s", acct.masked())
        # 预检老号令牌，失效则立即换下一个（避免向读不了的老号发码、空等超时）
        if not graph_mail.token_alive(acct.refresh_token, proxy_url=read_proxy):
            logger.warning("恢复老号 %s 令牌失效(invalid_grant)，跳过", acct.masked())
            continue
        since_iso = _iso_now_minus(60)

        # ---- 1) AddProof：把老号设为恢复邮箱，触发发码 ----
        add_data = {
            "iProofOptions": "Email",
            "DisplayPhoneCountryISO": country,
            "DisplayPhoneNumber": "",
            "EmailAddress": acct.email,
            "canary": add_canary,
            "action": "AddProof",
            "PhoneNumber": "",
            "PhoneCountryISO": "",
        }
        try:
            r_add = http.post(
                add_action, data=add_data,
                headers={
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Origin": "https://account.live.com",
                    "Referer": resp.url or add_action,
                },
                allow_redirects=True,
            )
        except Exception as exc:  # noqa: BLE001
            logger.warning("AddProof 请求异常(%s): %s", acct.masked(), exc)
            continue
        _dump_html("proofs_add_resp.html", r_add.text or "")
        verify_form = _parse_form_by_id(r_add.text or "", "frmVerifyProof")
        if not verify_form:
            logger.warning(
                "AddProof 后未到 Verify 页(%s) status=%s url=%s（换下一个老号）",
                acct.masked(), r_add.status_code, (r_add.url or "")[:100],
            )
            continue
        logger.info("AddProof 成功，已到 proofs/Verify 页，等待老号收 OTT…")

        # ---- 2) 读 OTT（老号 refresh_token 走 Graph，time-filter 防旧码）----
        code = graph_mail.read_security_code(
            acct.refresh_token, mode="graph", timeout=150,
            proxy_url=read_proxy, since_iso=since_iso,
        )
        if not code:
            logger.warning("老号 %s 未读到 OTT（超时/被拒），换下一个老号", acct.masked())
            continue
        logger.info("已从恢复老号读到 OTT=%s", code)

        # ---- 3) VerifyProof：回填 OTT 完成验证 ----
        # 重新解析（AddProof 落地即 Verify 页）的表单，拿最新 canary + proof 描述符
        verify_form = _parse_form_by_id(r_add.text or "", "frmVerifyProof") or verify_form
        verify_action = urllib.parse.urljoin(r_add.url or "", verify_form["action"])
        vf = verify_form["fields"]
        proof_opt = ""
        for _k, v in vf.items():
            if v.startswith("OTT||"):
                proof_opt = v
                break
        if not proof_opt:
            proof_opt = f"OTT||{acct.email}||Email||0||t"
        verify_data = {
            "iProofOptions": proof_opt,
            "iOttText": code,
            "action": "VerifyProof",
            "canary": vf.get("canary", ""),
            "GeneralVerify": "0",
        }
        logger.info("VerifyProof 提交 iOttText=%s proof=%s canary=%s…",
                    code, proof_opt[:30], (vf.get("canary", "") or "")[:14])
        try:
            r_vf = http.post(
                verify_action, data=verify_data,
                headers={
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Origin": "https://account.live.com",
                    "Referer": r_add.url or verify_action,
                },
                allow_redirects=True,
            )
        except Exception as exc:  # noqa: BLE001
            logger.warning("VerifyProof 请求异常(%s): %s", acct.masked(), exc)
            continue
        _dump_html("proofs_verify_resp.html", r_vf.text or "")

        # ---- 4) 成功判定 + 续跳：成功页会带 frmSubmitSLT.slt（HAR 坐实），POST 它续登录 ----
        slt_val, slt_action, slt_fields = _extract_success_slt(r_vf.text or "", r_vf.url or "")
        if not slt_val:
            logger.warning(
                "VerifyProof 未获 slt（OTT=%s 可能被拒/canary 过期），换下一个老号；resp still_verify=%s url=%s",
                code, "frmVerifyProof" in (r_vf.text or ""), (r_vf.url or "")[:90],
            )
            continue
        logger.info("VerifyProof 成功（恢复老号 %s，已获续登录 slt）", acct.masked())
        r_slt = http.post(slt_action, data=slt_fields, allow_redirects=True)
        r_next = follow_auto_post_forms(
            http, r_slt, tag="afterproof", enable_proof_pool=False, max_hops=12, country=country,
        )
        nxt_url = (r_next.url or "").lower()
        if "code=" in nxt_url or "outlook.live.com" in nxt_url or not _is_proofs_add_page(r_next.text or "", nxt_url):
            logger.info("proofs 满足成功（恢复老号 %s）→ 续跳 url=%s",
                        acct.masked(), (r_next.url or "")[:110])
            meta = {
                "proofs_method": "outlook_pool",
                "recovery_email": acct.email,
                "recovery_password": acct.password,
                "proofs_satisfied": "true",
            }
            return r_next, meta
        logger.warning("VerifyProof 成功但续跳后仍在 proofs（%s），换下一个老号", acct.masked())

    logger.error("收码池所有老号均未能满足 proofs")
    return None


def _do_proof_round(
    http: OutlookHttpSession,
    *,
    add_action: str,
    add_canary: str,
    referer: str,
    recovery_email: str,
    country: str,
    read_code,
    log_label: str,
    tag: str = "afterproof_ext",
) -> Optional[requests.Response]:
    """单个恢复邮箱的 AddProof → 收 OTT → VerifyProof → SLT 续登录（收码后端无关）。

    ``read_code``：无参回调，在 AddProof 成功（微软已发码）后调用，返回验证码或空串。
    IMAP 后端里读 IMAP，CF 后端里读 CF Worker——本函数只负责微软侧协议，不关心收码来源。
    成功返回续跳后的 response；任一步失败返回 None（调用方换下一个恢复邮箱）。
    """
    add_data = {
        "iProofOptions": "Email",
        "DisplayPhoneCountryISO": country,
        "DisplayPhoneNumber": "",
        "EmailAddress": recovery_email,
        "canary": add_canary,
        "action": "AddProof",
        "PhoneNumber": "",
        "PhoneCountryISO": "",
    }
    try:
        r_add = http.post(
            add_action, data=add_data,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Origin": "https://account.live.com",
                "Referer": referer or add_action,
            },
            allow_redirects=True,
        )
    except Exception as exc:  # noqa: BLE001
        logger.warning("AddProof 异常(%s): %s", log_label, exc)
        return None
    verify_form = _parse_form_by_id(r_add.text or "", "frmVerifyProof")
    if not verify_form:
        logger.warning("AddProof 后未到 Verify 页(%s)，换下一个", log_label)
        return None
    code = read_code()
    if not code:
        logger.warning("恢复邮箱 %s 未读到 OTT，换下一个", log_label)
        return None
    verify_form = _parse_form_by_id(r_add.text or "", "frmVerifyProof") or verify_form
    verify_action = urllib.parse.urljoin(r_add.url or "", verify_form["action"])
    vf = verify_form["fields"]
    proof_opt = next(
        (v for _k, v in vf.items() if v.startswith("OTT||")),
        f"OTT||{recovery_email}||Email||0||t",
    )
    verify_data = {
        "iProofOptions": proof_opt,
        "iOttText": code,
        "action": "VerifyProof",
        "canary": vf.get("canary", ""),
        "GeneralVerify": "0",
    }
    try:
        r_vf = http.post(
            verify_action, data=verify_data,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Origin": "https://account.live.com",
                "Referer": r_add.url or verify_action,
            },
            allow_redirects=True,
        )
    except Exception as exc:  # noqa: BLE001
        logger.warning("VerifyProof 异常(%s): %s", log_label, exc)
        return None
    slt_val, slt_action, slt_fields = _extract_success_slt(r_vf.text or "", r_vf.url or "")
    if not slt_val:
        return None
    r_slt = http.post(slt_action, data=slt_fields, allow_redirects=True)
    r_next = follow_auto_post_forms(
        http, r_slt, tag=tag, enable_proof_pool=False, max_hops=12, country=country,
    )
    nxt_url = (r_next.url or "").lower()
    if "code=" in nxt_url or "outlook.live.com" in nxt_url or not _is_proofs_add_page(r_next.text or "", nxt_url):
        return r_next
    logger.warning("VerifyProof 成功但续跳后仍在 proofs（%s），换下一个", log_label)
    return None


def satisfy_proofs_with_external(
    http: OutlookHttpSession,
    resp: requests.Response,
    *,
    max_accounts: int = 6,
    country: str = "US",
) -> Optional[tuple[requests.Response, dict[str, str]]]:
    """撞到 proofs/Add 时用**外部恢复邮箱**满足 proofs（AddProof + Verify）。

    收码后端由 OUTLOOK_RECOVERY_BACKEND 切换：
      imap（默认） — login.exe 同款第三方 IMAP 恢复邮箱池（your-recovery-host.com 等）。
      cf_domain    — Cloudflare 域名 catch-all 邮箱：按需生成 xxxx@域名，经 CF Worker API 收码。
      coolhs_mail  — 自建 coolhs-mail（hook.coolhs.com）：按需生成 @mail.coolhs.com，经 HTTP API 收码。
    """
    from . import external_recovery_pool as ext_pool

    if not ext_pool.external_pool_enabled():
        return None

    body = resp.text or ""
    add_form = _parse_form_by_id(body, "frmAddProof")
    if not add_form:
        logger.warning("proofs 页未解析到 frmAddProof，无法用外部恢复邮箱")
        return None
    add_action = urllib.parse.urljoin(resp.url or "", add_form["action"])
    add_canary = add_form["fields"].get("canary", "")
    referer = resp.url or add_action

    backend = ext_pool.recovery_backend()
    if backend == "cf_domain":
        return _satisfy_proofs_cf_domain(
            http, add_action=add_action, add_canary=add_canary, referer=referer,
            max_accounts=max_accounts, country=country,
        )
    if backend == "coolhs_mail":
        return _satisfy_proofs_coolhs_mail(
            http, add_action=add_action, add_canary=add_canary, referer=referer,
            max_accounts=max_accounts, country=country,
        )
    return _satisfy_proofs_imap(
        http, add_action=add_action, add_canary=add_canary, referer=referer,
        max_accounts=max_accounts, country=country,
    )


def _satisfy_proofs_imap(
    http: OutlookHttpSession,
    *,
    add_action: str,
    add_canary: str,
    referer: str,
    max_accounts: int,
    country: str,
) -> Optional[tuple[requests.Response, dict[str, str]]]:
    """IMAP 恢复邮箱池后端（原有路径，行为不变）。"""
    from . import external_recovery_pool as ext_pool
    from .mail_reader import read_security_code_imap_password

    host = ext_pool.imap_host()
    port = ext_pool.imap_port()
    since_ts = time.time() - 60

    for acct in ext_pool.iter_accounts(limit=max_accounts):
        logger.info("proofs 外部恢复邮箱(IMAP)：选用 %s", acct.masked())

        def _read(a=acct):
            return read_security_code_imap_password(
                a.email, a.password,
                imap_host=host, imap_port=port, since_ts=since_ts, timeout=150,
            )

        r_next = _do_proof_round(
            http, add_action=add_action, add_canary=add_canary, referer=referer,
            recovery_email=acct.email, country=country, read_code=_read,
            log_label=acct.masked(), tag="afterproof_ext",
        )
        if r_next is not None:
            meta = {
                "proofs_method": "external_recovery",
                "recovery_email": acct.email,
                "recovery_password": acct.password,
                "proofs_satisfied": "true",
            }
            logger.info("proofs 外部恢复邮箱(IMAP)绑定成功 → %s", acct.masked())
            return r_next, meta
    logger.error("外部恢复邮箱池(IMAP)均未能满足 proofs")
    return None


def _satisfy_proofs_cf_domain(
    http: OutlookHttpSession,
    *,
    add_action: str,
    add_canary: str,
    referer: str,
    max_accounts: int,
    country: str,
) -> Optional[tuple[requests.Response, dict[str, str]]]:
    """Cloudflare 域名 catch-all 邮箱后端：按需生成地址 + CF Worker API 收码。"""
    from . import cf_domain_mail

    try:
        client = cf_domain_mail.CFDomainMailClient()
    except Exception as exc:  # noqa: BLE001
        logger.error("CF 域名邮箱后端初始化失败：%s", exc)
        return None

    placeholder = cf_domain_mail.recovery_placeholder()
    try:
        max_try = int(os.environ.get("OUTLOOK_CF_PROOF_MAX_ACCOUNTS", str(max_accounts)))
    except ValueError:
        max_try = max_accounts
    for _i in range(max(1, max_try)):
        try:
            address = cf_domain_mail.allocate_address(client)
        except Exception as exc:  # noqa: BLE001
            logger.error("CF 域名邮箱分配地址失败：%s", exc)
            return None
        logger.info("proofs 外部恢复邮箱(CF域名)：分配 %s", address)

        try:
            before_ids = client.snapshot_ids(address)
        except Exception:  # noqa: BLE001
            before_ids = set()

        def _read(addr=address, bids=before_ids):
            # AddProof 已成功、微软已发码后再计时
            return client.read_security_code(
                addr, since_ts=time.time() - 30, before_ids=bids, timeout=150,
            )

        r_next = _do_proof_round(
            http, add_action=add_action, add_canary=add_canary, referer=referer,
            recovery_email=address, country=country, read_code=_read,
            log_label=address, tag="afterproof_cf",
        )
        if r_next is not None:
            meta = {
                "proofs_method": "cf_domain_recovery",
                "recovery_email": address,
                "recovery_password": placeholder,
                "proofs_satisfied": "true",
            }
            logger.info("proofs 外部恢复邮箱(CF域名)绑定成功 → %s", address)
            return r_next, meta
    logger.error("CF 域名恢复邮箱均未能满足 proofs（%d 次尝试）", max_accounts)
    return None


def _satisfy_proofs_coolhs_mail(
    http: OutlookHttpSession,
    *,
    add_action: str,
    add_canary: str,
    referer: str,
    max_accounts: int,
    country: str,
) -> Optional[tuple[requests.Response, dict[str, str]]]:
    """coolhs-mail 恢复邮箱后端：按需生成 @mail.coolhs.com + HTTP API 收码。"""
    from . import coolhs_mail

    try:
        client = coolhs_mail.CoolhsMailClient()
    except Exception as exc:  # noqa: BLE001
        logger.error("coolhs-mail 后端初始化失败：%s", exc)
        return None

    placeholder = coolhs_mail.recovery_placeholder()
    try:
        max_try = int(os.environ.get("OUTLOOK_COOLHS_PROOF_MAX_ACCOUNTS", str(max_accounts)))
    except ValueError:
        max_try = max_accounts
    for _i in range(max(1, max_try)):
        try:
            address = coolhs_mail.allocate_address(client)
        except Exception as exc:  # noqa: BLE001
            logger.error("coolhs-mail 分配地址失败：%s", exc)
            return None
        logger.info("proofs 外部恢复邮箱(coolhs-mail)：分配 %s", address)

        try:
            before_ids = client.snapshot_ids(address)
        except Exception:  # noqa: BLE001
            before_ids = set()

        def _read(addr=address, bids=before_ids):
            return client.read_security_code(
                addr, since_ts=time.time() - 30, before_ids=bids, timeout=150,
            )

        r_next = _do_proof_round(
            http, add_action=add_action, add_canary=add_canary, referer=referer,
            recovery_email=address, country=country, read_code=_read,
            log_label=address, tag="afterproof_coolhs",
        )
        if r_next is not None:
            meta = {
                "proofs_method": "coolhs_mail_recovery",
                "recovery_email": address,
                "recovery_password": placeholder,
                "proofs_satisfied": "true",
            }
            logger.info("proofs 外部恢复邮箱(coolhs-mail)绑定成功 → %s", address)
            return r_next, meta
    logger.error("coolhs-mail 恢复邮箱均未能满足 proofs（%d 次尝试）", max_accounts)
    return None


def _dump_html(name: str, text: str) -> None:
    if not os.environ.get("OUTLOOK_OAUTH_DEBUG"):
        return
    d = Path("debug_oauth")
    d.mkdir(exist_ok=True)
    (d / name).write_text(text, encoding="utf-8", errors="replace")
    logger.info("已 dump %s (%d bytes)", name, len(text))


def follow_auto_post_forms(
    http: OutlookHttpSession,
    resp: requests.Response,
    *,
    max_hops: int = 6,
    tag: str = "",
    enable_proof_pool: Optional[bool] = None,
    proof_meta: Optional[dict[str, str]] = None,
    country: str = "US",
    ctx: Optional[SignupSession] = None,
) -> requests.Response:
    """跟随 MSA 的 JS 自动提交隐藏表单（fmHF 等），直到落到最终 URL。

    requests 不执行 JS，MSA 登录链路里大量用 `<form>` + `window.onload submit`，
    必须手动解析表单字段并逐跳 POST。
    """
    from .proof_pool import pool_enabled

    if enable_proof_pool is None:
        enable_proof_pool = pool_enabled()
    proof_tried = False
    seen_urls: list[str] = []
    for hop in range(max_hops):
        body = resp.text or ""
        cur_url = resp.url or ""
        _dump_html(f"{tag}_hop{hop}.html", body)
        if "code=" in cur_url:
            return resp

        # 0a) OAuth 同意授权确认页 → POST ucaction=Yes（消费者账号首次授权 Thunderbird 客户端必经）
        consent_resp = submit_consent(http, resp)
        if consent_resp is not None:
            resp = consent_resp
            continue

        # 0b) KMSI「保持登录?」插页 → 续跳
        kmsi_resp = submit_kmsi(http, resp)
        if kmsi_resp is not None:
            resp = kmsi_resp
            continue

        # 0c) credentialaction 路由插页 → passkey / proofs / OAuth（US 新号常见）
        if _is_credentialaction_interrupt(body, cur_url):
            ca_resp = try_handle_credentialaction(http, resp, ctx)
            if ca_resp is not None:
                resp = ca_resp
                continue

        # 0d) Passkey 插页 → HAR: enroll error_code=NotAllowedError（禁止 POST fido/create）
        if _is_passkey_interrupt(body, cur_url):
            skip_resp = try_skip_passkey(http, body, ctx, url=cur_url)
            if skip_resp is not None:
                resp = skip_resp
                continue

        # 1) 账号安全信息插页（proofs/Add）
        #    优先：外部恢复邮箱（login.exe 同款 your-recovery-host.com）
        #    次选：Outlook 收码池老号
        #    回退：仅 OUTLOOK_SKIP_PROOFS=1 时 cancel 跳过
        if "account.live.com" in cur_url or "/proofs" in cur_url.lower():
            if not proof_tried and _is_proofs_add_page(body, cur_url):
                proof_tried = True
                ext = satisfy_proofs_with_external(http, resp, country=country)
                if ext is not None:
                    resp, meta = ext
                    if proof_meta is not None:
                        proof_meta.update(meta)
                    continue
                if enable_proof_pool:
                    pool_ext = satisfy_proofs_with_pool(http, resp, country=country)
                    if pool_ext is not None:
                        resp, meta = pool_ext
                        if proof_meta is not None:
                            proof_meta.update(meta)
                        continue
                if _skip_proofs_allowed():
                    logger.info("收码池未满足 proofs，允许 skip → cancel 跳过")
                else:
                    logger.error(
                        "proofs 未完成且 OUTLOOK_SKIP_PROOFS 未开启，停止 cancel 跳过"
                        "（请配置 OUTLOOK_EXTERNAL_RECOVERY_POOL_FILE + OUTLOOK_RECOVERY_IMAP_HOST）"
                    )
                    return resp
            if _skip_proofs_allowed():
                skip = _find_skip_url(body)
                if skip and skip not in seen_urls:
                    seen_urls.append(skip)
                    logger.info("跳过账号安全信息插页 → %s", skip[:100])
                    resp = http.get(urllib.parse.urljoin(cur_url, skip), allow_redirects=True)
                    continue

        # 2) 静态自动提交表单（fmHF 等）
        form_m = _FORM_RE.search(body)
        if form_m:
            inner = form_m.group(1)
            # 守卫：proofs 需人工输入的表单（frmVerifyProof/frmAddProof）不可盲目重提，
            # 否则空 iOttText 重提会 302 到 error.aspx?errcode=1086。改提交成功续跳的 frmSubmitSLT。
            if re.search(r'name=["\'](iOttText|EmailAddress)["\']', inner):
                slt, slt_action, slt_fields = _extract_success_slt(body, cur_url)
                if slt:
                    resp = http.post(slt_action, data=slt_fields, allow_redirects=True)
                    continue
                logger.debug("遇到 proofs 人工输入表单但无续跳 slt，交由 proofs 分支/返回")
                return resp
            form_open = body[form_m.start():body.find(">", form_m.start()) + 1]
            action = urllib.parse.urljoin(cur_url, _attr(form_open, "action") or cur_url)
            if "credentialaction" in action.lower() and _is_credentialaction_interrupt(body, cur_url):
                ca_resp = try_handle_credentialaction(http, resp, ctx)
                if ca_resp is not None:
                    resp = ca_resp
                    continue
            if "fido/create" in action.lower():
                skip_resp = try_skip_passkey(http, body, ctx, url=cur_url)
                if skip_resp is not None:
                    resp = skip_resp
                    continue
                logger.warning("Passkey fido/create 表单无法取消，停止跟随以免误报名")
                return resp
            method = (_attr(form_open, "method") or "post").lower()
            fields = {}
            for inp in _INPUT_RE.findall(form_m.group(1)):
                n = _attr(inp, "name")
                if n:
                    fields[n] = _attr(inp, "value")
            # 说明：早期用 posturl=ppsecure/post.srf「桥接」跳过加验证，但实测该 POST
            # 返回登录页 sErrorCode=80041032 死循环。正解是把 fmHF 正常 POST 到 proofs/Add
            # 加载出真正的安全信息页（其 $Config viewDefs.cancel.url 即「跳过」续跳到
            # oauth20_authorize.srf），下一跳由 _find_skip_url 命中该 cancel URL GET 跳过。
            logger.debug("自动表单跳转 hop=%s method=%s action=%s fields=%s",
                         hop, method, action[:100], list(fields.keys()))
            if method == "get":
                resp = http.get(action, params=fields, allow_redirects=True)
            else:
                resp = http.post(action, data=fields, allow_redirects=True)
            if "credentialaction" in action.lower():
                ca_resp = try_handle_credentialaction(http, resp, ctx, post_fields=fields)
                if ca_resp is not None:
                    resp = ca_resp
            continue

        # 3) AAD BssoInterrupt/$Config 驱动（无静态 form）→ POST urlPost 续跳
        url_post = _config_str(body, "urlPost")
        if url_post and cur_url not in seen_urls:
            seen_urls.append(cur_url)
            action = urllib.parse.urljoin(cur_url, url_post)
            canary = _config_str(body, "canary")
            canary_name = _config_str(body, "sCanaryTokenName") or "canary"
            ft = _config_str(body, "sFT")
            ft_name = _config_str(body, "sFTName") or "flowToken"
            fields = {}
            if canary:
                fields[canary_name] = canary
            if ft:
                fields[ft_name] = ft
            logger.debug("BSSO 续跳 hop=%s action=%s fields=%s", hop, action[:100], list(fields.keys()))
            resp = http.post(action, data=fields, allow_redirects=True)
            continue

        return resp
    logger.warning("自动表单跳转超过 %s 跳仍未完成 tag=%s url=%s", max_hops, tag, (resp.url or "")[:120])
    return resp


def skip_proofs_interstitial(http: OutlookHttpSession, resp: requests.Response, *, max_hops: int = 4) -> requests.Response:
    """跳过「添加安全信息」插页：GET cancel viewDef 的 oauth 续跳地址。

    fresh 账号常被要求加手机/备用邮箱；页面 hasCancel=1，GET cancel.url 即可续跳。
    """
    for _ in range(max_hops):
        body = resp.text or ""
        cur = resp.url or ""
        on_proofs = ("account.live.com" in cur and "proof" in cur.lower()) or "AddProofControl" in body
        if not on_proofs:
            return resp
        skip = _find_skip_url(body)
        if not skip:
            logger.warning("proofs 插页未找到 cancel 链接，无法跳过")
            return resp
        logger.info("跳过账号安全信息插页 → %s", skip[:100])
        resp = http.get(urllib.parse.urljoin(cur, skip), allow_redirects=True)
        resp = follow_auto_post_forms(http, resp, tag="proofskip")
    return resp


def submit_slt_login(
    http: OutlookHttpSession,
    ctx: SignupSession,
    redirect_url: str,
    slt: str,
    *,
    proof_meta: Optional[dict[str, str]] = None,
    country: str = "US",
) -> requests.Response:
    resp = http.post(
        redirect_url,
        data={"slt": slt},
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": "https://signup.live.com",
            "Referer": ctx.signup_page_url,
        },
        allow_redirects=True,
    )
    resp = follow_auto_post_forms(
        http, resp, tag="slt", proof_meta=proof_meta, country=country, ctx=ctx,
    )
    if _skip_proofs_allowed():
        resp = skip_proofs_interstitial(http, resp)
    logger.info("slt 登录完成 status=%s url=%s proofs=%s",
                resp.status_code, (resp.url or "")[:120],
                (proof_meta or {}).get("proofs_method", "none"))
    return resp


def _is_credentialaction_interrupt(body: str, url: str = "") -> bool:
    """MSA 凭证动作路由插页（2025+ US 等新号常见）。

    抓包/日志证据（captcha_run_e2e_20260814_171208.log）：
      slt/oauth fmHF → POST account.live.com/interrupt/credentialaction
      字段 scenarios, mpcxt, pprid, ipt, posturl, uaid
      响应 ~11KB SPA，无 fmHF；旧逻辑在此停住，OAuth 拿不到 code。
    """
    low = (url or "").lower()
    if "interrupt/credentialaction" in low:
        return True
    blob = body or ""
    return ("interrupt/credentialaction" in blob.lower()) and ("scenarios" in blob.lower())


def _collect_hidden_fields(body: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    for inp in _INPUT_RE.findall(body or ""):
        n = _attr(inp, "name")
        if n and n not in fields:
            fields[n] = _attr(inp, "value")
    return fields


def _find_account_interrupt_urls(body: str, base_url: str) -> list[str]:
    """从 account.live.com 插页 HTML/$Config 中提取 proofs/passkey 等续跳 URL。"""
    found: list[str] = []
    seen: set[str] = set()
    patterns = [
        r'"(https://account\.live\.com/(?:interrupt|proofs)/[^"]+)"',
        r'"(/interrupt/[^"]+)"',
        r'"(/proofs/[^"]+)"',
        r'href="(https://account\.live\.com/(?:interrupt|proofs)/[^"]+)"',
        r'action="(https://account\.live\.com/(?:interrupt|proofs)/[^"]+)"',
    ]
    for pat in patterns:
        for m in re.finditer(pat, body or "", re.I):
            u = html_lib.unescape(m.group(1)).replace("&amp;", "&")
            if not u.startswith("http"):
                u = urllib.parse.urljoin(base_url, u)
            if u not in seen:
                seen.add(u)
                found.append(u)
    return found


def _credentialaction_scenarios(body: str, fields: dict[str, str]) -> str:
    if fields.get("scenarios"):
        return fields["scenarios"]
    m = re.search(r'"scenarios"\s*:\s*"((?:[^"\\]|\\.)*)"', body or "")
    return _decode_js_str(m.group(1)) if m else ""


def try_handle_credentialaction(
    http: OutlookHttpSession,
    resp: requests.Response,
    ctx: Optional[SignupSession] = None,
    *,
    post_fields: Optional[dict[str, str]] = None,
) -> Optional[requests.Response]:
    """credentialaction 路由页续跳 → passkey 取消 / proofs/Add / OAuth。

    HAR 成功链（Outlook抓包.har）无 credentialaction（旧区域直连 proofs/Add）；
    失败日志坐实 credentialaction 为 passkey+proofs 的上游路由器。
    """
    body = resp.text or ""
    url = resp.url or ""
    if not _is_credentialaction_interrupt(body, url):
        return None

    fields = _collect_hidden_fields(body)
    if post_fields:
        for k, v in post_fields.items():
            if v and k not in fields:
                fields[k] = v
    scenarios = _credentialaction_scenarios(body, fields)
    logger.info(
        "credentialaction 插页续跳 scenarios=%s hidden=%s",
        (scenarios or "")[:120], list(fields.keys()),
    )

    url_post = _config_str(body, "urlPost")
    if url_post:
        action = urllib.parse.urljoin(url, url_post)
        post_fields: dict[str, str] = {}
        canary = _config_str(body, "canary") or fields.get("canary", "")
        if canary:
            post_fields["canary"] = canary
        if scenarios:
            post_fields["scenarios"] = scenarios
        for key in ("pprid", "ipt", "uaid", "mpcxt", "posturl"):
            if fields.get(key):
                post_fields[key] = fields[key]
        logger.info("credentialaction $Config.urlPost → %s", action[:100])
        nxt = http.post(action, data=post_fields, allow_redirects=True)
        if not _is_credentialaction_interrupt(nxt.text or "", nxt.url or ""):
            return nxt

    for target in _find_account_interrupt_urls(body, url):
        low = target.lower()
        if "passkey" in low:
            logger.info("credentialaction 页内 passkey 链 → %s", target[:100])
            data = {k: fields[k] for k in ("pprid", "ipt", "uaid") if fields.get(k)}
            r = http.post(target, data=data, allow_redirects=True) if data else http.get(target, allow_redirects=True)
            skipped = try_skip_passkey(http, r.text or "", ctx, url=r.url or target)
            r = skipped or r
            if not _is_credentialaction_interrupt(r.text or "", r.url or ""):
                return r
        elif "proofs" in low:
            logger.info("credentialaction 页内 proofs 链 → %s", target[:100])
            data = {k: fields[k] for k in ("pprid", "ipt", "uaid") if fields.get(k)}
            r = http.post(target, data=data, allow_redirects=True) if data else http.get(target, allow_redirects=True)
            if not _is_credentialaction_interrupt(r.text or "", r.url or ""):
                return r

    parsed = urllib.parse.urlparse(url)
    qs = urllib.parse.parse_qs(parsed.query)
    ru = qs.get("ru", [""])[0]
    mkt = qs.get("mkt", ["EN-US"])[0]
    client_id = qs.get("client_id", ["1E0000487A244A"])[0]
    cobrandid = qs.get("cobrandid", [""])[0]
    page_id = qs.get("id", ["292841"])[0]
    if not (fields.get("pprid") and fields.get("ipt")):
        logger.warning("credentialaction 缺 pprid/ipt，无法构造 passkey/proofs 续跳")
        return None

    passkey_qs = {
        "mkt": mkt,
        "uiflavor": qs.get("uiflavor", ["web"])[0],
        "client_id": client_id,
        "id": page_id,
        "fluent": "2",
        "ru": ru,
    }
    if cobrandid:
        passkey_qs["cobrandid"] = cobrandid
    passkey_url = "https://account.live.com/interrupt/passkey?" + urllib.parse.urlencode(passkey_qs)
    passkey_data = {k: fields[k] for k in ("pprid", "ipt", "uaid") if fields.get(k)}
    logger.info("credentialaction 构造 POST passkey（对齐 Outlook抓包.har idx 100）")
    r_pk = http.post(passkey_url, data=passkey_data, allow_redirects=True)
    skipped = try_skip_passkey(http, r_pk.text or "", ctx, url=r_pk.url or passkey_url)
    r_pk = skipped or r_pk
    if not _is_credentialaction_interrupt(r_pk.text or "", r_pk.url or ""):
        return r_pk

    proofs_qs = {
        "mkt": mkt,
        "uiflavor": qs.get("uiflavor", ["web"])[0],
        "client_id": client_id,
        "id": page_id,
        "mpcxt": fields.get("mpcxt", "CATB"),
        "ru": ru,
    }
    if fields.get("posturl"):
        proofs_qs["posturl"] = fields["posturl"]
    if cobrandid:
        proofs_qs["cobrandid"] = cobrandid
    proofs_url = "https://account.live.com/proofs/Add?" + urllib.parse.urlencode(proofs_qs)
    proofs_data = {k: fields[k] for k in ("pprid", "ipt", "uaid") if fields.get(k)}
    logger.info("credentialaction 构造 POST proofs/Add（对齐 slt_hop0 fmHF → proofs/Add）")
    return http.post(proofs_url, data=proofs_data, allow_redirects=True)


def _is_passkey_interrupt(body: str, url: str = "") -> bool:
    low = (url or "").lower()
    if "interrupt/passkey" in low or "fido/create" in low:
        return True
    blob = body or ""
    if "passkey/enroll" in blob.lower():
        return True
    if "postBackUrl" in blob and "passkey" in blob.lower():
        return True
    return False


def try_skip_passkey(
    http: OutlookHttpSession,
    html: str,
    ctx: Optional[SignupSession] = None,
    *,
    url: str = "",
) -> Optional[requests.Response]:
    """取消 Passkey 报名（对齐 Outlook抓包.har idx 120）。

    浏览器会把 fmHF POST 到 interrupt/passkey，页面再自动提交到
    login.microsoft.com/consumers/fido/create。纯 API 无法完成 WebAuthn；
    HAR 成功链在用户取消后 POST interrupt/passkey/enroll，
    error_code=NotAllowedError → 302 oauth20_authorize.srf?PasskeyEnrollResult=user_cancel。
    禁止把 fido/create 表单当真提交。
    """
    if not html or not _is_passkey_interrupt(html, url):
        return None

    fields: dict[str, str] = {}
    for inp in _INPUT_RE.findall(html):
        n = _attr(inp, "name")
        if n and n not in fields:
            fields[n] = _attr(inp, "value")
    post_back = fields.get("postBackUrl") or _config_str(html, "postBackUrl")
    canary = fields.get("canary") or _config_str(html, "canary") or _config_str(html, "sCanary")
    if not post_back or "passkey" not in post_back.lower():
        skip_patterns = [
            r'href="([^"]*skip[^"]*)"',
            r'"urlSkip"[^>]*value="([^"]+)"',
        ]
        for pat in skip_patterns:
            m = re.search(pat, html, re.I)
            if m:
                skip_url = m.group(1).replace("&amp;", "&")
                if not skip_url.startswith("http"):
                    skip_url = f"{ACCOUNT_BASE}{skip_url}"
                logger.info("Passkey 回退跳过链接: %s", skip_url[:100])
                return http.get(skip_url, allow_redirects=True)
        return None
    if not post_back.startswith("http"):
        post_back = urllib.parse.urljoin(ACCOUNT_BASE + "/", post_back)

    data = {
        "canary": canary,
        "authenticator": "",
        "transports": "",
        "aaguid": "",
        "credentialDeviceType": "",
        "credentialBackedUp": "",
        "attestationParseError": "",
        "error_code": "NotAllowedError",
        "suberror_code": "",
        "error_message": (
            "The operation either timed out or was not allowed. "
            "See: https://www.w3.org/TR/webauthn-2/#sctn-privacy-considerations-client."
        ),
        "mediation": "",
        "clientDataJson": "",
        "attestationObject": "",
        "credentialId": "",
        "clientExtensionResults": "",
        "i19": "",
    }
    logger.info("Passkey 按 HAR 取消报名 enroll error_code=NotAllowedError")
    return http.post(
        post_back,
        data=data,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": "https://account.live.com",
            "Referer": post_back.split("?")[0],
        },
        allow_redirects=True,
    )


def fetch_mail_oauth_code(
    http: OutlookHttpSession,
    ctx: SignupSession,
    email: str,
    *,
    client_id: str = MAIL_CLIENT_ID,
    scope: str = MAIL_SCOPE,
    proof_meta: Optional[dict[str, str]] = None,
) -> dict[str, str]:
    """注册后获取邮件 OAuth code（exe 同款 client_id）。需在 slt 登录完成后调用。

    关键：proofs+slt 登录后会话建立在 **login.live.com（消费者 MSA）**，而非
    login.microsoftonline.com（AAD）。用 AAD /common/authorize 会因无 ESTS 会话落到
    ConvergedSignIn 登录页（AADSTS900144）。参考工具（Outlook抓包.har）对 9e5f94bc 直接
    GET login.live.com/oauth20_authorize.srf，复用消费者会话 → Consent/Update(ucaction=Yes)
    → nativeclient?code。故优先走消费者端点，失败再回退 AAD 端点。
    """
    base_params = {
        "client_id": client_id,
        "scope": scope,
        "redirect_uri": MAIL_REDIRECT_URI,
        "response_type": "code",
        "response_mode": "query",
        "login_hint": email,
        "uaid": ctx.uaid,
        "msproxy": "1",
        "issuer": "mso",
        "tenant": "common",
        "ui_locales": ctx.mkt.replace("_", "-") if ctx.mkt else "en-US",
    }
    endpoints = [
        ("consumer", "https://login.live.com/oauth20_authorize.srf"),
        ("aad", "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"),
    ]
    info: dict[str, str] = {}
    for name, ep in endpoints:
        url = f"{ep}?{urllib.parse.urlencode(base_params)}"
        logger.info("读信授权（%s）…", name)
        resp = http.get(url, allow_redirects=True)
        resp = follow_auto_post_forms(
            http, resp, tag=f"authorize_{name}", max_hops=12, proof_meta=proof_meta,
        )
        final = resp.url or ""
        info["authorize_url"] = final
        m = re.search(r"[?&]code=([^&]+)", final)
        if m:
            info["code"] = urllib.parse.unquote(m.group(1))
            logger.info("获取读信授权成功（%s）", name)
            return info
        _dump_html(f"authorize_{name}_final.html", resp.text or "")
        logger.warning("（%s）未从跳转解析到授权码", name)
    return info


def exchange_code_for_token(
    http: OutlookHttpSession,
    code: str,
    *,
    client_id: str = MAIL_CLIENT_ID,
    scope: str = MAIL_SCOPE,
) -> dict[str, str]:
    """authorization_code → refresh_token（公共 client，无 secret）。"""
    resp = http.post(
        "https://login.microsoftonline.com/common/oauth2/v2.0/token",
        data={
            "client_id": client_id,
            "code": code,
            "redirect_uri": MAIL_REDIRECT_URI,
            "grant_type": "authorization_code",
            "scope": scope,
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        allow_redirects=False,
    )
    out: dict[str, str] = {}
    try:
        data = resp.json()
    except ValueError:
        logger.warning("token 交换响应非 JSON status=%s body=%s", resp.status_code, resp.text[:200])
        return out
    if resp.status_code >= 400 or data.get("error"):
        logger.warning(
            "token 交换失败 status=%s error=%s desc=%s",
            resp.status_code, data.get("error"), (data.get("error_description") or "")[:160],
        )
        return out
    for key in ("refresh_token", "access_token", "expires_in", "scope", "token_type", "id_token"):
        if data.get(key) is not None:
            out[key] = str(data[key])
    if out.get("refresh_token"):
        logger.info("换取令牌成功，已获取读信令牌")
    elif out.get("id_token") or out.get("access_token"):
        # 拿到 id_token/access_token 但没 refresh_token → 多半是 scope 缺 offline_access
        logger.warning("换取令牌只返回部分结果，未拿到读信令牌")
    return out


def fetch_login_token(
    http: OutlookHttpSession,
    ctx: SignupSession,
    email: str,
    *,
    attempts: int = LOGIN_TOKEN_ATTEMPTS,
) -> dict[str, str]:
    """双令牌 token#2（登录授权）：authorize→code→token，带重试与显式失败原因。

    返回 dict 含 refresh_token/id_token/access_token/scope（成功时），以及：
      login_status: ok | no_code | no_refresh | error
      login_fail_reason: 失败原因（error/error_description 或阶段）
    每次尝试都会重走一遍 authorize（消费者会话仍在，通常直接回 code）。
    """
    out: dict[str, str] = {}
    last_reason = ""
    for i in range(1, max(1, attempts) + 1):
        code_info = fetch_mail_oauth_code(
            http, ctx, email, client_id=LOGIN_CLIENT_ID, scope=LOGIN_SCOPE,
        )
        # 记录 authorize 落点，便于诊断被 consent/interrupt 打断的情况
        if code_info.get("authorize_url"):
            out["authorize_url"] = code_info["authorize_url"]
        code = code_info.get("code", "")
        if not code:
            last_reason = f"authorize 未拿到 code(attempt={i}) url={code_info.get('authorize_url','')[:120]}"
            logger.warning("token#2 %s", last_reason)
            out["status"] = "no_code"
            out["fail_reason"] = last_reason
            time.sleep(1.5)
            continue
        ltok = exchange_code_for_token(
            http, code, client_id=LOGIN_CLIENT_ID, scope=LOGIN_SCOPE,
        )
        for k, v in ltok.items():
            out[k] = v
        if ltok.get("refresh_token"):
            out["status"] = "ok"
            out.pop("fail_reason", None)
            logger.info("登录令牌获取成功")
            logger.debug(
                "token#2 成功(attempt=%d): refresh_token=True id_token=%s scope=%s",
                i, bool(ltok.get("id_token")), ltok.get("scope", ""),
            )
            return out
        # 有 code 但换不到 refresh_token（只回 id_token/access_token 或直接报错）
        if ltok.get("id_token") or ltok.get("access_token"):
            last_reason = f"仅返回 id_token/access_token 无 refresh_token(attempt={i}) scope={ltok.get('scope','')}"
            out["status"] = "no_refresh"
        else:
            last_reason = f"token 交换失败(attempt={i})"
            out["status"] = "error"
        out["fail_reason"] = last_reason
        logger.warning("token#2 %s，重试…", last_reason)
        time.sleep(1.5)
    logger.error("token#2 最终失败(%d 次)：%s", attempts, last_reason)
    return out


def complete_oauth_after_signup(
    http: OutlookHttpSession,
    ctx: SignupSession,
    redirect_url: str,
    slt: str,
    *,
    email: str = "",
    password: str = "",
    proxy: Optional[str] = None,
    fetch_mail_token: bool = False,
    country: str = "US",
) -> dict[str, str]:
    proof_meta: dict[str, str] = {}
    resp = submit_slt_login(
        http, ctx, redirect_url, slt, proof_meta=proof_meta, country=country,
    )
    final_url = resp.url or ""
    logged_in = "outlook.live.com" in final_url or "login.srf" in final_url or "code=" in final_url
    info: dict[str, str] = {"final_url": final_url, "logged_in": str(logged_in)}
    info.update(proof_meta)

    # 仍停在 credentialaction / passkey 插页时再续跳一次
    if _is_credentialaction_interrupt(resp.text or "", final_url):
        ca = try_handle_credentialaction(http, resp, ctx)
        info["credentialaction_handled"] = str(ca is not None)
        if ca is not None:
            resp = follow_auto_post_forms(
                http, ca, tag="after_credaction", proof_meta=proof_meta, country=country, ctx=ctx,
            )
            final_url = resp.url or ""
            info["final_url"] = final_url
            logged_in = (
                "outlook.live.com" in final_url
                or "login.srf" in final_url
                or "code=" in final_url
            )
            info["logged_in"] = str(logged_in)

    if _is_passkey_interrupt(resp.text or "", final_url):
        skipped = try_skip_passkey(http, resp.text or "", ctx, url=final_url)
        info["passkey_skipped"] = str(skipped is not None)
        if skipped is not None:
            resp = follow_auto_post_forms(
                http, skipped, tag="after_passkey", proof_meta=proof_meta, country=country, ctx=ctx,
            )
            final_url = resp.url or ""
            info["final_url"] = final_url
            logged_in = (
                "outlook.live.com" in final_url
                or "login.srf" in final_url
                or "code=" in final_url
            )
            info["logged_in"] = str(logged_in)

    if fetch_mail_token and email:
        # 快路径：纯 requests
        mail_info = fetch_mail_oauth_code(http, ctx, email)
        info.update({f"mail_{k}": v for k, v in mail_info.items()})
        if mail_info.get("code"):
            token_info = exchange_code_for_token(http, mail_info["code"])
            info.update({f"mail_{k}": v for k, v in token_info.items()})

        # 双令牌 token#2：登录授权（第三方 SSO）。复用已建立的消费者会话，
        # 用 LOGIN_CLIENT_ID + LOGIN_SCOPE 再做一次 authorize→code→token 交换（带重试）。
        if DUAL_TOKEN and info.get("mail_refresh_token"):
            try:
                login_info = fetch_login_token(http, ctx, email)
                info.update({f"login_{k}": v for k, v in login_info.items()})
            except Exception as exc:  # noqa: BLE001
                logger.warning("双令牌 token#2 获取异常: %s", exc)
                info["login_status"] = "error"
                info["login_fail_reason"] = str(exc)
            if not info.get("login_refresh_token"):
                logger.warning("双令牌最终未产出读信令牌，本号回落只存四段")

    logger.info("注册后登录链路完成: logged_in=%s mail_refresh=%s login_refresh=%s",
                info.get("logged_in"), bool(info.get("mail_refresh_token")),
                bool(info.get("login_refresh_token")))
    return info
