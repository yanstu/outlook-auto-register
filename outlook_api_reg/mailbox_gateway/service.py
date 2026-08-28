"""邮箱解析、字段裁剪、读信 / 取码。

读信一律复用 `graph_mail`：先自己用 refresh_token 换一次 access_token（顺带判活、
拿到微软轮换出来的新 refresh_token 并写回账号库），再让 graph_mail 拉列表。
四段令牌本身不出库、不出接口。
"""
from __future__ import annotations

import base64
import hmac
import logging
import time
from datetime import datetime
from typing import Any, Optional

from .. import account_store, graph_mail, lifecycle
from ..constants import MAIL_CLIENT_ID
from .errors import MailboxApiError
from .store import normalize_scopes

logger = logging.getLogger(__name__)

MAILBOX_ID_PREFIX = "mbx_"
MAX_WAIT_SECONDS = 60
OTP_POLL_INTERVAL = 5.0
FOLDERS = ("inbox", "junkemail")

_SECURITY_SUBJECT_HINTS = (
    "security code", "verification code", "verify", "安全代码", "验证码",
    "验证你的", "single-use code", "one-time",
)


# ── mailbox_id ──────────────────────────────────────────────────────────────


def mailbox_id_for(email: str) -> str:
    raw = (email or "").strip().lower().encode("utf-8")
    return MAILBOX_ID_PREFIX + base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def email_from_mailbox_id(mailbox_id: str) -> str:
    text = (mailbox_id or "").strip()
    if not text.startswith(MAILBOX_ID_PREFIX):
        return ""
    body = text[len(MAILBOX_ID_PREFIX):]
    pad = "=" * (-len(body) % 4)
    try:
        return base64.urlsafe_b64decode(body + pad).decode("utf-8").strip().lower()
    except Exception:  # noqa: BLE001
        return ""


def parse_mailbox_ref(ref: str) -> str:
    """mailbox_id 或邮箱地址 → 邮箱地址。"""
    ref = (ref or "").strip()
    if "@" in ref:
        return ref.lower()
    return email_from_mailbox_id(ref)


# ── 邮箱资源 ────────────────────────────────────────────────────────────────


def _readable(row: dict[str, Any]) -> bool:
    verify = row.get("verify") or {}
    return bool(verify.get("ok") or verify.get("graph"))


def mailbox_view(row: dict[str, Any], scopes: list[str]) -> dict[str, Any]:
    """按 scope 裁剪出一个邮箱的可见字段。"""
    scopes = normalize_scopes(scopes)
    email = (row.get("email") or "").strip()
    view: dict[str, Any] = {
        "mailbox_id": mailbox_id_for(email),
        "email": email,
    }
    if "fields:basic" in scopes or "mailboxes:read" in scopes:
        view.update({
            "batch": row.get("batch_label") or "",
            "batch_no": row.get("batch_no"),
            "created_at": row.get("created_at") or "",
            "incubating": bool(row.get("incubating")),
            "incubation_until": row.get("incubation_until") or "",
            "readable": _readable(row),
            "has_token": bool(row.get("has_token")),
            "last_alive_at": row.get("last_alive_at") or "",
            "tags": list(row.get("tags") or []),
        })
    if "fields:sensitive" in scopes:
        view.update({
            "password": row.get("password") or "",
            "client_id": row.get("client_id") or "",
            "recovery_email": row.get("recovery_email") or "",
            "recovery_password": row.get("recovery_password") or "",
        })
    return view


def combo_line(row: dict[str, Any]) -> str:
    """六段优先、四段兜底的 combo 文本。"""
    six = lifecycle.combo_recovery_line(row)
    if six:
        return six
    combo = (row.get("combo") or "").strip()
    if combo:
        return combo
    email = (row.get("email") or "").strip()
    if not email:
        return ""
    return "----".join([
        email,
        row.get("password") or "",
        row.get("client_id") or "",
        row.get("refresh_token") or "",
    ])


def fields_view(row: dict[str, Any], profile: str, scopes: list[str]) -> dict[str, Any]:
    scopes = normalize_scopes(scopes)
    profile = (profile or "basic").strip().lower()
    email = (row.get("email") or "").strip()
    if profile not in ("basic", "full", "combo"):
        raise MailboxApiError(400, "bad_profile", "profile 只能是 basic / full / combo")
    if "fields:basic" not in scopes:
        raise MailboxApiError(403, "scope_required", "当前令牌缺少 fields:basic 权限",
                              required_scope="fields:basic")
    if profile in ("full", "combo") and "fields:sensitive" not in scopes:
        raise MailboxApiError(403, "scope_required",
                              f"profile={profile} 需要 fields:sensitive 权限",
                              required_scope="fields:sensitive")
    basic = {
        "email": email,
        "readable": _readable(row),
        "incubating": bool(row.get("incubating")),
        "batch": row.get("batch_label") or "",
        "created_at": row.get("created_at") or "",
    }
    if profile == "basic":
        return {"mailbox_id": mailbox_id_for(email), "profile": profile, "fields": basic}
    if profile == "full":
        full = dict(basic)
        full.update({
            "password": row.get("password") or "",
            "recovery_email": row.get("recovery_email") or "",
            "recovery_password": row.get("recovery_password") or "",
            "client_id": row.get("client_id") or "",
        })
        return {"mailbox_id": mailbox_id_for(email), "profile": profile, "fields": full}
    line = combo_line(row)
    return {
        "mailbox_id": mailbox_id_for(email),
        "profile": profile,
        "combo": line,
        "segments": len(line.split("----")) if line else 0,
    }


def load_mailbox(email: str) -> dict[str, Any]:
    row = account_store.get_account(email)
    if not row:
        raise MailboxApiError(404, "mailbox_not_found", "没有这个邮箱", email=email)
    lifecycle.enrich_lifecycle_fields(row)
    return row


def list_mailboxes(
    principal: Any,
    *,
    limit: int = 50,
    offset: int = 0,
    q: str = "",
    batch: str = "",
    readable_only: bool = False,
) -> dict[str, Any]:
    rows = account_store.list_accounts()
    q = (q or "").strip().lower()
    batch = (batch or "").strip().lower()
    picked: list[dict[str, Any]] = []
    for row in rows:
        email = (row.get("email") or "").strip()
        if not email or not principal.can_access(email):
            continue
        if q and q not in email.lower():
            continue
        if batch and batch != (row.get("batch_label") or "").strip().lower():
            continue
        if readable_only and not _readable(row):
            continue
        picked.append(row)
    total = len(picked)
    window = picked[max(0, offset): max(0, offset) + max(1, min(limit, 500))]
    return {
        "total": total,
        "limit": limit,
        "offset": offset,
        "items": [mailbox_view(r, principal.scopes_for(r.get("email") or "")) for r in window],
    }


def readable_count(principal: Any) -> int:
    n = 0
    for row in account_store.list_accounts():
        email = (row.get("email") or "").strip()
        if email and principal.can_access(email) and _readable(row):
            n += 1
    return n


# ── 读信前置：判活 + 令牌轮换写回 ────────────────────────────────────────────


def ensure_not_incubating(row: dict[str, Any], principal: Any) -> None:
    if principal.has_scope("incubation:bypass"):
        return
    if not row.get("incubating"):
        return
    raise MailboxApiError(
        423,
        "incubating",
        "账号还在孵化期，暂不开放读信",
        until=row.get("incubation_until") or "",
        email=(row.get("email") or "").strip(),
    )


def _writeback_rotation(email: str, old_rt: str, new_rt: str) -> None:
    """微软轮换出新 refresh_token 时同步回账号库（combo 各段一并跟上）。"""
    if not new_rt or new_rt == old_rt:
        return
    row = account_store.get_account(email) or {}
    now = datetime.now().isoformat(timespec="seconds")
    patch: dict[str, Any] = {
        "refresh_token": new_rt,
        "updated_at": now,
        "last_alive_at": now,
    }
    for key in ("combo", "combo_dual", "combo_recovery"):
        line = (row.get(key) or "").strip()
        parts = line.split("----")
        if len(parts) >= 4:
            parts[3] = new_rt
            patch[key] = "----".join(parts)
    try:
        account_store.patch_account(email, patch)
    except Exception as exc:  # noqa: BLE001
        logger.warning("轮换后的读信令牌写回失败 %s: %s", email, exc)


def prepare_token(row: dict[str, Any], mode: str = "auto") -> tuple[str, str]:
    """→ (可用的 refresh_token, 实际读信通道)。顺带把轮换出的新令牌写回库。"""
    email = (row.get("email") or "").strip()
    rt = (row.get("refresh_token") or "").strip()
    if not rt:
        raise MailboxApiError(404, "no_token", "这个邮箱没有读信令牌", email=email)
    mode = (mode or "auto").strip().lower()
    if mode not in ("auto", "graph", "outlook_rest"):
        raise MailboxApiError(400, "bad_mode", "mode 只能是 auto / graph / outlook_rest")

    client_id = (row.get("client_id") or "").strip() or MAIL_CLIENT_ID
    data = graph_mail.refresh_token_for(rt, "", client_id=client_id)
    if data.get("transient"):
        raise MailboxApiError(503, "upstream_unavailable", "微软接口暂时连不上，请稍后重试",
                              email=email)
    if not data.get("access_token"):
        raise MailboxApiError(
            502,
            "token_dead",
            "读信令牌已失效，需要重登或救援",
            email=email,
            detail=str(data.get("error_description") or data.get("error") or "")[:120],
        )
    new_rt = (data.get("refresh_token") or "").strip()
    if new_rt and new_rt != rt:
        _writeback_rotation(email, rt, new_rt)
        rt = new_rt
    if mode != "auto":
        return rt, mode
    resolved = "outlook_rest" if graph_mail._resource_of(data.get("scope", "")) == "outlook" else "graph"
    return rt, resolved


# ── 读信 ────────────────────────────────────────────────────────────────────


def _message_view(m: dict[str, Any], folder: str) -> dict[str, Any]:
    return {
        "id": m.get("id") or "",
        "folder": folder,
        "subject": m.get("subject") or "",
        "from": m.get("from") or "",
        "received": m.get("received") or "",
        "preview": m.get("preview") or "",
        "body": m.get("body") or "",
    }


def _matches(m: dict[str, Any], *, after: str, sender: str, subject_contains: str) -> bool:
    if after and (m.get("received") or "") < after:
        return False
    if sender and sender.lower() not in (m.get("from") or "").lower():
        return False
    if subject_contains and subject_contains.lower() not in (m.get("subject") or "").lower():
        return False
    return True


def list_messages(
    row: dict[str, Any],
    *,
    folder: str = "inbox",
    limit: int = 20,
    after: str = "",
    sender: str = "",
    subject_contains: str = "",
    mode: str = "auto",
) -> dict[str, Any]:
    folder = (folder or "inbox").strip().lower()
    if folder not in FOLDERS:
        raise MailboxApiError(400, "bad_folder", "folder 只能是 inbox 或 junkemail")
    limit = max(1, min(int(limit or 20), 50))
    rt, resolved = prepare_token(row, mode)
    raw = graph_mail.list_messages(rt, mode=resolved, top=limit, folder=folder)
    items = [
        _message_view(m, folder)
        for m in raw
        if _matches(m, after=after, sender=sender, subject_contains=subject_contains)
    ]
    return {
        "email": (row.get("email") or "").strip(),
        "folder": folder,
        "mode": resolved,
        "count": len(items),
        "messages": items[:limit],
    }


def get_message(row: dict[str, Any], message_id: str, *, mode: str = "auto") -> dict[str, Any]:
    message_id = (message_id or "").strip()
    if not message_id:
        raise MailboxApiError(400, "bad_message_id", "缺少 message_id")
    rt, resolved = prepare_token(row, mode)
    for folder in FOLDERS:
        for m in graph_mail.list_messages(rt, mode=resolved, top=50, folder=folder):
            if (m.get("id") or "") == message_id:
                out = _message_view(m, folder)
                out["email"] = (row.get("email") or "").strip()
                out["mode"] = resolved
                return out
    raise MailboxApiError(404, "message_not_found", "最近的收件箱 / 垃圾邮件里没有这封信",
                          email=(row.get("email") or "").strip())


def wait_otp(
    row: dict[str, Any],
    *,
    wait_seconds: int = 60,
    after: str = "",
    sender: str = "",
    subject_contains: str = "",
    mode: str = "auto",
) -> dict[str, Any]:
    """长轮询取验证码：命中即返回，超时返回 found=false。"""
    wait_seconds = max(0, min(int(wait_seconds or 0), MAX_WAIT_SECONDS))
    rt, resolved = prepare_token(row, mode)
    email = (row.get("email") or "").strip()
    deadline = time.monotonic() + wait_seconds
    strict = bool(sender or subject_contains)
    while True:
        for folder in FOLDERS:
            for m in graph_mail.list_messages(rt, mode=resolved, top=15, folder=folder):
                if not _matches(m, after=after, sender=sender, subject_contains=subject_contains):
                    continue
                if not strict and not _looks_like_otp_mail(m):
                    continue
                code = graph_mail._extract_code(
                    m.get("subject") or "", m.get("body") or "", m.get("preview") or ""
                )
                if code:
                    return {
                        "email": email,
                        "found": True,
                        "code": code,
                        "mode": resolved,
                        "message": _message_view(m, folder),
                    }
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        time.sleep(min(OTP_POLL_INTERVAL, remaining))
    return {"email": email, "found": False, "code": "", "mode": resolved,
            "waited_seconds": wait_seconds}


def _looks_like_otp_mail(m: dict[str, Any]) -> bool:
    sender = (m.get("from") or "").lower()
    subject = (m.get("subject") or "").lower()
    if any(s in sender for s in graph_mail._SECURITY_SENDERS):
        return True
    return any(h in subject for h in _SECURITY_SUBJECT_HINTS)


def verify_account_password(email: str, password: str) -> Optional[dict[str, Any]]:
    """用户端登录：邮箱 + 邮箱密码。密码不匹配返回 None。"""
    email = (email or "").strip().lower()
    if not email or not password:
        return None
    row = account_store.get_account(email)
    if not row:
        return None
    stored = (row.get("password") or "").strip()
    if not stored or not hmac.compare_digest(stored, password):
        return None
    lifecycle.enrich_lifecycle_fields(row)
    return row
