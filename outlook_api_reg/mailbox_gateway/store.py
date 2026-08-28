"""Mailbox API 的凭据存储：principal / grant / session / audit 四张表的读写。

令牌格式统一为 `<前缀><principal_id>_<secret>`：
principal_id 是 12 位十六进制，用来做一次索引命中；secret 是 32 字节随机串，
只以 pbkdf2-sha256 摘要落库，明文仅在创建那一刻返回给调用方。
"""
from __future__ import annotations

import hashlib
import hmac
import os
import secrets
import threading
from datetime import datetime, timedelta
from typing import Any, Optional

from .. import database as db

KEY_PREFIX = "mbx_sk_"
SESSION_PREFIX = "mbx_sess_"
BOOTSTRAP_PRINCIPAL_ID = "bootstrap"
BOOTSTRAP_SCOPES = "admin,*"

PBKDF2_ITERATIONS = 120_000
_ALGO = "pbkdf2_sha256"

DEFAULT_SESSION_HOURS = 24
USER_SESSION_SCOPES = ("fields:basic", "messages:read", "otp:read")

_lock = threading.Lock()


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _new_id() -> str:
    return secrets.token_hex(6)


def _new_secret() -> str:
    return secrets.token_urlsafe(32)


# ── 摘要 ────────────────────────────────────────────────────────────────────


def hash_secret(secret: str, *, iterations: int = PBKDF2_ITERATIONS) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", secret.encode("utf-8"), salt, iterations)
    return f"{_ALGO}${iterations}${salt.hex()}${digest.hex()}"


def verify_secret(secret: str, stored: str) -> bool:
    if not secret or not stored:
        return False
    try:
        algo, iters, salt_hex, digest_hex = stored.split("$", 3)
    except ValueError:
        return False
    if algo != _ALGO:
        return False
    try:
        digest = hashlib.pbkdf2_hmac(
            "sha256", secret.encode("utf-8"), bytes.fromhex(salt_hex), int(iters)
        )
    except (ValueError, TypeError):
        return False
    return hmac.compare_digest(digest.hex(), digest_hex)


def normalize_scopes(raw: Any) -> list[str]:
    """逗号分隔 / 列表 → 去重有序的小写 scope 列表。"""
    if raw is None:
        return []
    items = raw if isinstance(raw, (list, tuple, set)) else str(raw).split(",")
    out: list[str] = []
    for item in items:
        s = str(item).strip().lower()
        if s and s not in out:
            out.append(s)
    return out


def scopes_text(raw: Any) -> str:
    return ",".join(normalize_scopes(raw))


def split_token(token: str, prefix: str) -> tuple[str, str]:
    """`<prefix><id>_<secret>` → (id, secret)；不合格式返回 ('', '')。"""
    if not token or not token.startswith(prefix):
        return "", ""
    rest = token[len(prefix):]
    pid, sep, secret = rest.partition("_")
    if not sep or not pid or not secret:
        return "", ""
    return pid, secret


# ── principal ───────────────────────────────────────────────────────────────


def _principal_row(row: Any) -> dict[str, Any]:
    d = dict(row)
    d["enabled"] = bool(d.get("enabled"))
    d["scopes"] = normalize_scopes(d.get("scopes"))
    return d


def get_principal(principal_id: str) -> Optional[dict[str, Any]]:
    if not principal_id:
        return None
    db.ensure_initialized()
    conn = db.connect()
    try:
        row = conn.execute("SELECT * FROM api_principals WHERE id=?", (principal_id,)).fetchone()
        return _principal_row(row) if row else None
    finally:
        conn.close()


def list_principals(*, kind: str = "") -> list[dict[str, Any]]:
    db.ensure_initialized()
    conn = db.connect()
    try:
        if kind:
            rows = conn.execute(
                "SELECT * FROM api_principals WHERE kind=? ORDER BY created_at DESC", (kind,)
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM api_principals ORDER BY created_at DESC").fetchall()
        out = []
        for row in rows:
            item = _principal_row(row)
            item.pop("secret_hash", None)
            item["grants"] = [g["email"] for g in grants_for(item["id"])]
            out.append(item)
        return out
    finally:
        conn.close()


def grants_for(principal_id: str) -> list[dict[str, Any]]:
    db.ensure_initialized()
    conn = db.connect()
    try:
        rows = conn.execute(
            "SELECT email, scopes_override FROM api_grants WHERE principal_id=?", (principal_id,)
        ).fetchall()
        return [
            {"email": (r["email"] or "").strip().lower(), "scopes_override": r["scopes_override"] or ""}
            for r in rows
        ]
    finally:
        conn.close()


def create_service_key(
    name: str,
    scopes: Any,
    *,
    grants: Optional[list[str]] = None,
    expires_at: str = "",
) -> dict[str, Any]:
    """建一把 service key。返回 {id, key, name, scopes, grants}，key 只在此刻可见。"""
    principal_id = _new_id()
    secret = _new_secret()
    scope_list = normalize_scopes(scopes)
    emails = [e.strip().lower() for e in (grants or []) if (e or "").strip()]
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            conn.execute(
                """INSERT INTO api_principals(id, kind, name, secret_hash, scopes, enabled, created_at, expires_at)
                   VALUES (?,?,?,?,?,1,?,?)""",
                (
                    principal_id,
                    "service",
                    (name or "").strip() or principal_id,
                    hash_secret(secret),
                    ",".join(scope_list),
                    _now(),
                    (expires_at or "").strip(),
                ),
            )
            for email in emails or [""]:
                conn.execute(
                    """INSERT OR IGNORE INTO api_grants(principal_id, email, scopes_override, created_at)
                       VALUES (?,?,?,?)""",
                    (principal_id, email, "", _now()),
                )
            conn.commit()
        finally:
            conn.close()
    return {
        "id": principal_id,
        "key": f"{KEY_PREFIX}{principal_id}_{secret}",
        "name": (name or "").strip() or principal_id,
        "scopes": scope_list,
        "grants": emails,
    }


def revoke_principal(principal_id: str) -> bool:
    """停用一个 principal 并清掉它名下的会话。"""
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            cur = conn.execute("UPDATE api_principals SET enabled=0 WHERE id=?", (principal_id,))
            conn.execute("DELETE FROM api_sessions WHERE principal_id=?", (principal_id,))
            conn.commit()
            return cur.rowcount > 0
        finally:
            conn.close()


def ensure_user_principal(email: str) -> dict[str, Any]:
    """用户端登录用的 principal：一个邮箱一个，固定只授单邮箱的只读 scope。"""
    email = (email or "").strip().lower()
    principal_id = "usr" + hashlib.sha256(email.encode("utf-8")).hexdigest()[:9]
    scopes = ",".join(USER_SESSION_SCOPES)
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            conn.execute(
                """INSERT INTO api_principals(id, kind, name, secret_hash, scopes, enabled, created_at, expires_at)
                   VALUES (?,?,?,'',?,1,?,'')
                   ON CONFLICT(id) DO UPDATE SET scopes=excluded.scopes, enabled=1""",
                (principal_id, "user", email, scopes, _now()),
            )
            conn.execute(
                """INSERT OR IGNORE INTO api_grants(principal_id, email, scopes_override, created_at)
                   VALUES (?,?,?,?)""",
                (principal_id, email, scopes, _now()),
            )
            conn.commit()
        finally:
            conn.close()
    return {
        "id": principal_id,
        "kind": "user",
        "name": email,
        "scopes": list(USER_SESSION_SCOPES),
        "enabled": True,
    }


# ── session ─────────────────────────────────────────────────────────────────


def session_hours() -> int:
    raw = (os.environ.get("OUTLOOK_MAILBOX_API_SESSION_HOURS") or "").strip()
    try:
        hours = int(float(raw)) if raw else DEFAULT_SESSION_HOURS
    except ValueError:
        hours = DEFAULT_SESSION_HOURS
    return max(1, min(hours, 24 * 30))


def create_session(principal_id: str, email: str, *, hours: Optional[int] = None) -> dict[str, Any]:
    session_id = _new_id()
    secret = _new_secret()
    h = session_hours() if hours is None else max(1, int(hours))
    expires = datetime.now() + timedelta(hours=h)
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            conn.execute(
                """INSERT INTO api_sessions(id, token_hash, principal_id, email, expires_at, created_at)
                   VALUES (?,?,?,?,?,?)""",
                (
                    session_id,
                    hash_secret(secret),
                    principal_id,
                    (email or "").strip().lower(),
                    expires.isoformat(timespec="seconds"),
                    _now(),
                ),
            )
            conn.commit()
        finally:
            conn.close()
    return {
        "id": session_id,
        "token": f"{SESSION_PREFIX}{session_id}_{secret}",
        "expires_at": expires.isoformat(timespec="seconds"),
        "expires_in": h * 3600,
    }


def get_session(session_id: str) -> Optional[dict[str, Any]]:
    if not session_id:
        return None
    db.ensure_initialized()
    conn = db.connect()
    try:
        row = conn.execute("SELECT * FROM api_sessions WHERE id=?", (session_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def delete_session(session_id: str) -> None:
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            conn.execute("DELETE FROM api_sessions WHERE id=?", (session_id,))
            conn.commit()
        finally:
            conn.close()


def purge_expired_sessions() -> int:
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            cur = conn.execute(
                "DELETE FROM api_sessions WHERE expires_at!='' AND expires_at < ?", (_now(),)
            )
            conn.commit()
            return cur.rowcount
        finally:
            conn.close()


# ── audit ───────────────────────────────────────────────────────────────────


def write_audit(
    *,
    principal_id: str,
    method: str,
    path: str,
    email: str = "",
    status: int = 200,
    detail: str = "",
) -> None:
    """一行简要访问记录。只写路径与结果，不写任何令牌 / 密码。"""
    try:
        db.ensure_initialized()
        conn = db.connect()
        try:
            conn.execute(
                """INSERT INTO api_audit(ts, principal_id, method, path, email, status, detail)
                   VALUES (?,?,?,?,?,?,?)""",
                (
                    _now(),
                    principal_id or "",
                    method or "",
                    (path or "")[:200],
                    (email or "").strip().lower(),
                    int(status or 0),
                    (detail or "")[:240],
                ),
            )
            conn.commit()
        finally:
            conn.close()
    except Exception:  # noqa: BLE001 — 审计写失败不该影响业务响应
        pass


def recent_audit(limit: int = 100) -> list[dict[str, Any]]:
    db.ensure_initialized()
    conn = db.connect()
    try:
        rows = conn.execute(
            "SELECT * FROM api_audit ORDER BY id DESC LIMIT ?", (max(1, min(limit, 1000)),)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ── bootstrap ───────────────────────────────────────────────────────────────


def bootstrap_admin_from_env() -> bool:
    """env OUTLOOK_MAILBOX_API_ADMIN_KEY 非空时，写入 / 刷新固定的 admin principal。"""
    raw = (os.environ.get("OUTLOOK_MAILBOX_API_ADMIN_KEY") or "").strip()
    if not raw:
        return False
    existing = get_principal(BOOTSTRAP_PRINCIPAL_ID)
    if existing and existing.get("enabled") and verify_secret(raw, existing.get("secret_hash") or ""):
        return False
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            conn.execute(
                """INSERT INTO api_principals(id, kind, name, secret_hash, scopes, enabled, created_at, expires_at)
                   VALUES (?,?,?,?,?,1,?,'')
                   ON CONFLICT(id) DO UPDATE SET
                     secret_hash=excluded.secret_hash,
                     scopes=excluded.scopes,
                     enabled=1""",
                (
                    BOOTSTRAP_PRINCIPAL_ID,
                    "service",
                    "bootstrap-admin",
                    hash_secret(raw),
                    BOOTSTRAP_SCOPES,
                    _now(),
                ),
            )
            conn.execute(
                """INSERT OR IGNORE INTO api_grants(principal_id, email, scopes_override, created_at)
                   VALUES (?,'','',?)""",
                (BOOTSTRAP_PRINCIPAL_ID, _now()),
            )
            conn.commit()
        finally:
            conn.close()
    return True
