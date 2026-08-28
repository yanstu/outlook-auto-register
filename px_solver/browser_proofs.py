#!/usr/bin/env python3
"""在比特/Playwright 当前页填写 proofs/Add + Verify，绑恢复邮箱。

OAuth 仍交给 ss_post.finish_after_proofs(proofs_done=True)。
"""
from __future__ import annotations

import os
import random
import sys
import time
from typing import Callable, Optional

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)


def browser_proofs_enabled() -> bool:
    raw = os.environ.get("BIT_BROWSER_PROOFS") or os.environ.get("SS_BROWSER_PROOFS") or "1"
    return raw.strip() != "0"


def classify_proofs_view(
    url: str = "",
    html: str = "",
    *,
    has_email: Optional[bool] = None,
    has_ott: Optional[bool] = None,
) -> str:
    """返回 add / verify / done / unknown。"""
    if has_ott:
        return "verify"
    if has_email:
        return "add"
    u = (url or "").lower()
    h = (html or "").lower()
    if "iotttext" in h or "frmverifyproof" in h or "/proofs/verify" in u:
        return "verify"
    if "emailaddress" in h or "frmaddproof" in h or "/proofs/add" in u:
        return "add"
    if any(k in u for k in (
        "privacynotice", "account.microsoft.com", "outlook.live.com",
        "oauth20", "/fp/",
    )):
        return "done"
    if "proof" in u:
        return "unknown"
    return "done"


def plan_recovery() -> Optional[dict]:
    """分配恢复邮箱（CF/coolhs 按需建址；IMAP 只读池文件）。"""
    from outlook_api_reg import external_recovery_pool as ext_pool

    if not ext_pool.external_pool_enabled():
        return None
    backend = ext_pool.recovery_backend()
    if backend == "cf_domain":
        from outlook_api_reg import cf_domain_mail

        client = cf_domain_mail.CFDomainMailClient()
        address = cf_domain_mail.allocate_address(client)
        return {
            "email": address,
            "password": cf_domain_mail.recovery_placeholder(),
            "method": "browser_proofs",
            "backend": "cf_domain",
            "client": client,
        }
    if backend == "coolhs_mail":
        from outlook_api_reg import coolhs_mail

        client = coolhs_mail.CoolhsMailClient()
        address = coolhs_mail.allocate_address(client)
        return {
            "email": address,
            "password": coolhs_mail.recovery_placeholder(),
            "method": "browser_proofs",
            "backend": "coolhs_mail",
            "client": client,
        }
    acct = next(iter(ext_pool.iter_accounts(limit=1)), None)
    if not acct:
        return None
    return {
        "email": acct.email,
        "password": acct.password,
        "method": "browser_proofs",
        "backend": "imap",
        "client": None,
    }


def _make_read_code(rec: dict, *, before_ids: Optional[set] = None, since_ts: float = 0.0) -> Callable[[], str]:
    if rec.get("backend") in ("cf_domain", "coolhs_mail") and rec.get("client") is not None:
        client = rec["client"]
        addr = rec["email"]
        bids = set(before_ids or [])

        def _read_http() -> str:
            return client.read_security_code(
                addr, since_ts=time.time() - 30, before_ids=bids, timeout=150,
            )

        return _read_http

    from outlook_api_reg import external_recovery_pool as ext_pool
    from outlook_api_reg.mail_reader import read_security_code_imap_password

    host = ext_pool.imap_host()
    port = ext_pool.imap_port()
    st = since_ts or (time.time() - 60)
    addr = rec["email"]
    pwd = rec["password"]

    def _read_imap() -> str:
        return read_security_code_imap_password(
            addr, pwd, imap_host=host, imap_port=port, since_ts=st, timeout=150,
        )

    return _read_imap


def _page_view(page) -> str:
    url = ""
    html = ""
    has_email = False
    has_ott = False
    try:
        url = page.url or ""
    except Exception:
        pass
    try:
        html = page.content() or ""
    except Exception:
        pass
    try:
        el = page.query_selector("#EmailAddress") or page.query_selector('input[name="EmailAddress"]')
        has_email = bool(el and el.bounding_box())
    except Exception:
        pass
    try:
        el = page.query_selector("#iOttText")
        has_ott = bool(el and el.bounding_box())
    except Exception:
        pass
    return classify_proofs_view(url, html, has_email=has_email, has_ott=has_ott)


def _click_proofs_next(page) -> bool:
    for sel in ("#iNext", 'input[id="iNext"]', 'input[type="submit"]', 'button[type="submit"]'):
        try:
            page.click(sel, timeout=3500)
            return True
        except Exception:
            continue
    for name in ("Next", "Verify", "Continue"):
        try:
            page.get_by_role("button", name=name).first.click(timeout=2500)
            return True
        except Exception:
            continue
    return False


def _wait_view(page, want: tuple[str, ...], timeout: float) -> str:
    deadline = time.time() + timeout
    last = "unknown"
    while time.time() < deadline:
        last = _page_view(page)
        if last in want:
            return last
        time.sleep(0.5)
    return last


def bind_recovery_in_browser(page, log: Callable[..., None]) -> dict:
    """在当前页走 Add → 收码 → Verify。成功后页面应已离开 proofs。"""
    rec = plan_recovery()
    if not rec:
        return {"ok": False, "note": "恢复邮箱未配置（coolhs_mail / CF / IMAP 池）"}

    out = {
        "ok": False,
        "recovery_email": rec["email"],
        "recovery_password": rec["password"],
        "proofs_method": rec["method"],
        "note": "",
    }
    log("  浏览器填恢复邮箱 %s", rec["email"])

    before_ids: set = set()
    if rec.get("client") is not None:
        try:
            before_ids = rec["client"].snapshot_ids(rec["email"]) or set()
        except Exception as exc:  # noqa: BLE001
            log("  收件快照失败（继续）: %s", exc)

    view = _wait_view(page, ("add",), 20)
    if view != "add":
        out["note"] = f"未等到 EmailAddress 表单（view={view}）"
        return out

    el = page.query_selector("#EmailAddress") or page.query_selector('input[name="EmailAddress"]')
    if not el:
        out["note"] = "找不到 #EmailAddress"
        return out
    try:
        el.click()
        el.fill("")
    except Exception:
        pass
    page.keyboard.type(rec["email"], delay=random.randint(40, 90))
    time.sleep(random.uniform(0.4, 0.8))
    if not _click_proofs_next(page):
        out["note"] = "恢复邮箱页点 Next 失败"
        return out

    view = _wait_view(page, ("verify", "done"), 30)
    if view == "done":
        out["ok"] = True
        log("  ✓ 提交恢复邮箱后已离开 proofs")
        return out
    if view != "verify":
        out["note"] = f"提交恢复邮箱后未到验证码页（view={view}）"
        return out

    log("  等待发到 %s 的验证码…", rec["email"])
    code = _make_read_code(rec, before_ids=before_ids)()
    if not code:
        out["note"] = "未读到 OTT 验证码"
        return out
    ott = page.query_selector("#iOttText")
    if not ott:
        out["note"] = "找不到 #iOttText"
        return out
    try:
        ott.click()
        ott.fill("")
    except Exception:
        pass
    page.keyboard.type(code, delay=random.randint(35, 80))
    time.sleep(random.uniform(0.3, 0.7))
    _click_proofs_next(page)

    deadline = time.time() + 35
    while time.time() < deadline:
        view = _page_view(page)
        if view == "done":
            out["ok"] = True
            log("  ✓ 浏览器已绑定恢复邮箱 %s", rec["email"])
            return out
        if view == "verify":
            _click_proofs_next(page)
        time.sleep(0.6)

    out["note"] = "验证码已提交但未离开 proofs 页"
    return out
