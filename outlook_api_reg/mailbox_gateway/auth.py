"""Bearer 解析、principal 还原、scope 与邮箱授权判定。

两层身份：
- service principal：`Authorization: Bearer mbx_sk_<id>_<secret>`，按 scopes 与 grants 授权；
- user session：`Authorization: Bearer mbx_sess_<id>_<secret>`，只能看登录时绑定的那一个邮箱。
"""
from __future__ import annotations

import functools
import inspect
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Optional

from fastapi import Request

from . import store
from .errors import MailboxApiError, forbidden, unauthorized

KNOWN_SCOPES = (
    "mailboxes:read",
    "fields:basic",
    "fields:sensitive",
    "messages:read",
    "otp:read",
    "admin",
    "incubation:bypass",
)
WILDCARD = "*"


@dataclass
class Principal:
    id: str
    kind: str
    name: str
    scopes: list[str] = field(default_factory=list)
    all_mailboxes: bool = False
    emails: set[str] = field(default_factory=set)
    scope_overrides: dict[str, list[str]] = field(default_factory=dict)
    session_id: str = ""
    session_email: str = ""
    expires_at: str = ""

    @property
    def wildcard(self) -> bool:
        return WILDCARD in self.scopes

    def has_scope(self, scope: str) -> bool:
        return self.wildcard or scope in self.scopes

    def scopes_for(self, email: str) -> list[str]:
        """某个邮箱上的有效 scope：principal scope 与该邮箱 override 取交集。"""
        base = list(KNOWN_SCOPES) if self.wildcard else list(self.scopes)
        override = self.scope_overrides.get((email or "").strip().lower())
        if override:
            return [s for s in base if s in override]
        return base

    def can_access(self, email: str) -> bool:
        email = (email or "").strip().lower()
        if not email:
            return False
        if self.session_email:
            return email == self.session_email
        return self.all_mailboxes or email in self.emails

    def to_public(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "kind": self.kind,
            "name": self.name,
            "scopes": list(self.scopes),
            "all_mailboxes": self.all_mailboxes,
            "mailboxes": sorted(self.emails) if not self.all_mailboxes else [],
            "session_email": self.session_email,
            "expires_at": self.expires_at,
        }


def _expired(expires_at: str) -> bool:
    text = (expires_at or "").strip()
    if not text:
        return False
    try:
        return datetime.fromisoformat(text) < datetime.now()
    except ValueError:
        return False


def _principal_from_row(row: dict[str, Any]) -> Principal:
    emails: set[str] = set()
    all_mailboxes = False
    overrides: dict[str, list[str]] = {}
    grants = store.grants_for(row["id"])
    if not grants:
        all_mailboxes = True
    for g in grants:
        email = g["email"]
        if not email:
            all_mailboxes = True
            continue
        emails.add(email)
        override = store.normalize_scopes(g.get("scopes_override"))
        if override:
            overrides[email] = override
    return Principal(
        id=row["id"],
        kind=row.get("kind") or "service",
        name=row.get("name") or "",
        scopes=list(row.get("scopes") or []),
        all_mailboxes=all_mailboxes,
        emails=emails,
        scope_overrides=overrides,
        expires_at=row.get("expires_at") or "",
    )


def bearer_token(request: Request) -> str:
    raw = request.headers.get("authorization") or ""
    scheme, _, token = raw.partition(" ")
    if scheme.lower() != "bearer":
        return ""
    return token.strip()


def resolve_token(token: str) -> Optional[Principal]:
    """令牌 → Principal。任何一步不过就返回 None（不区分原因，避免探测）。"""
    token = (token or "").strip()
    if not token:
        return None

    if token.startswith(store.SESSION_PREFIX):
        session_id, secret = store.split_token(token, store.SESSION_PREFIX)
        if not session_id:
            return None
        session = store.get_session(session_id)
        if not session or not store.verify_secret(secret, session.get("token_hash") or ""):
            return None
        if _expired(session.get("expires_at") or ""):
            store.delete_session(session_id)
            return None
        row = store.get_principal(session.get("principal_id") or "")
        if not row or not row.get("enabled") or _expired(row.get("expires_at") or ""):
            return None
        principal = _principal_from_row(row)
        principal.session_id = session_id
        principal.session_email = (session.get("email") or "").strip().lower()
        principal.expires_at = session.get("expires_at") or ""
        # 会话只认绑定的那一个邮箱，principal 上的 grants 不放大范围
        principal.all_mailboxes = False
        principal.emails = {principal.session_email} if principal.session_email else set()
        return principal

    if token.startswith(store.KEY_PREFIX):
        principal_id, secret = store.split_token(token, store.KEY_PREFIX)
        row = store.get_principal(principal_id) if principal_id else None
        if row and row.get("enabled") and not _expired(row.get("expires_at") or ""):
            if store.verify_secret(secret, row.get("secret_hash") or ""):
                return _principal_from_row(row)

    # bootstrap admin：env 里给的原样字符串，没有 id 段
    boot = store.get_principal(store.BOOTSTRAP_PRINCIPAL_ID)
    if boot and boot.get("enabled") and store.verify_secret(token, boot.get("secret_hash") or ""):
        return _principal_from_row(boot)
    return None


async def current_principal(request: Request) -> Principal:
    """FastAPI 依赖：拿不到有效令牌直接 401。"""
    principal = resolve_token(bearer_token(request))
    if principal is None:
        store.write_audit(
            principal_id="",
            method=request.method,
            path=request.url.path,
            status=401,
            detail="令牌无效",
        )
        raise unauthorized()
    request.state.principal = principal
    return principal


def require_scope(principal: Principal, scope: str, *, email: str = "") -> None:
    """校验 principal 在（可选的）某个邮箱上是否持有该 scope。"""
    allowed = principal.scopes_for(email) if email else (
        list(KNOWN_SCOPES) if principal.wildcard else list(principal.scopes)
    )
    if scope not in allowed:
        raise forbidden(f"当前令牌缺少 {scope} 权限", code="scope_required", required_scope=scope)


def audited(handler: Callable) -> Callable:
    """给路由套一层访问记录：谁、什么方法、什么路径、结果状态。"""

    @functools.wraps(handler)
    async def wrapper(*args: Any, **kwargs: Any) -> Any:
        request: Optional[Request] = kwargs.get("request")
        principal: Optional[Principal] = kwargs.get("principal")
        email = ""
        status = 200
        detail = ""
        try:
            result = handler(*args, **kwargs)
            if inspect.isawaitable(result):
                result = await result
            if isinstance(result, dict):
                email = str(result.get("email") or "")
            return result
        except MailboxApiError as exc:
            status, detail = exc.status_code, exc.code
            raise
        except Exception as exc:  # noqa: BLE001
            status, detail = 500, type(exc).__name__
            raise
        finally:
            if request is not None:
                store.write_audit(
                    principal_id=principal.id if principal else "",
                    method=request.method,
                    path=request.url.path,
                    email=email or str(kwargs.get("email") or ""),
                    status=status,
                    detail=detail,
                )

    return wrapper
