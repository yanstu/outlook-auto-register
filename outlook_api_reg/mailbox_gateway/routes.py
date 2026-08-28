"""Mailbox API v1 路由（前缀 /api/v1）。"""
from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Depends, Query, Request

from .. import database as db
from . import service, store
from .auth import Principal, audited, current_principal, require_scope
from .errors import MailboxApiError
from .schemas import (
    CreateKeyRequest,
    CreateKeyResponse,
    HealthResponse,
    LoginRequest,
    LoginResponse,
    MailboxListResponse,
    MessageItem,
    MessageListResponse,
    OtpResponse,
    PrincipalResponse,
)

router = APIRouter(prefix="/api/v1", tags=["mailbox"])


def _resolve(principal: Principal, ref: str, scope: str) -> dict[str, Any]:
    """mailbox_id / 邮箱 → 账号行，并校验访问范围与 scope。"""
    email = service.parse_mailbox_ref(ref)
    if not email:
        raise MailboxApiError(400, "bad_mailbox_id", "mailbox_id 无法解析")
    if not principal.can_access(email):
        raise MailboxApiError(403, "mailbox_forbidden", "当前令牌无权访问这个邮箱", email=email)
    require_scope(principal, scope, email=email)
    return service.load_mailbox(email)


# ── 鉴权 ────────────────────────────────────────────────────────────────────


@router.post("/auth/login", response_model=LoginResponse)
async def login(request: Request, body: LoginRequest) -> dict[str, Any]:
    """用户端登录：邮箱 + 邮箱密码换一枚只能看自己这一个邮箱的会话令牌。"""
    email = (body.email or "").strip().lower()
    row = service.verify_account_password(email, body.password or "")
    if row is None:
        store.write_audit(
            principal_id="", method=request.method, path=request.url.path,
            email=email, status=401, detail="登录失败",
        )
        raise MailboxApiError(401, "invalid_credentials", "邮箱或密码不对")
    principal = store.ensure_user_principal(email)
    session = store.create_session(principal["id"], email)
    store.write_audit(
        principal_id=principal["id"], method=request.method, path=request.url.path,
        email=email, status=200, detail="登录成功",
    )
    return {
        "token": session["token"],
        "expires_in": session["expires_in"],
        "expires_at": session["expires_at"],
        "mailbox_id": service.mailbox_id_for(email),
        "email": email,
        "scopes": list(store.USER_SESSION_SCOPES),
    }


@router.post("/auth/keys", response_model=CreateKeyResponse)
@audited
async def create_key(
    request: Request,
    body: CreateKeyRequest,
    principal: Principal = Depends(current_principal),
) -> dict[str, Any]:
    """建一把 service key。返回的 key 只出现这一次，之后库里只有摘要。"""
    require_scope(principal, "admin")
    scopes = store.normalize_scopes(body.scopes)
    if not scopes:
        raise MailboxApiError(400, "scopes_required", "至少要给一个 scope")
    return store.create_service_key(
        body.name, scopes, grants=body.grants, expires_at=body.expires_at
    )


@router.get("/auth/keys")
@audited
async def list_keys(
    request: Request,
    principal: Principal = Depends(current_principal),
) -> dict[str, Any]:
    require_scope(principal, "admin")
    return {"items": store.list_principals()}


@router.delete("/auth/keys/{key_id}")
@audited
async def revoke_key(
    key_id: str,
    request: Request,
    principal: Principal = Depends(current_principal),
) -> dict[str, Any]:
    require_scope(principal, "admin")
    if not store.revoke_principal(key_id):
        raise MailboxApiError(404, "key_not_found", "没有这把 key")
    return {"ok": True, "id": key_id}


@router.get("/auth/me", response_model=PrincipalResponse)
@audited
async def whoami(
    request: Request,
    principal: Principal = Depends(current_principal),
) -> dict[str, Any]:
    return principal.to_public()


# ── 邮箱资源 ────────────────────────────────────────────────────────────────


@router.get("/mailboxes", response_model=MailboxListResponse, response_model_exclude_none=True)
@audited
async def list_mailboxes(
    request: Request,
    principal: Principal = Depends(current_principal),
    limit: int = Query(50, ge=1, le=500),
    offset: int = Query(0, ge=0),
    q: str = Query("", description="邮箱地址模糊匹配"),
    batch: str = Query("", description="按批次标签过滤"),
    readable_only: bool = Query(False),
) -> dict[str, Any]:
    require_scope(principal, "mailboxes:read")
    return service.list_mailboxes(
        principal, limit=limit, offset=offset, q=q, batch=batch, readable_only=readable_only
    )


@router.get("/mailboxes/by-email/{email}", response_model_exclude_none=True)
@audited
async def mailbox_by_email(
    email: str,
    request: Request,
    principal: Principal = Depends(current_principal),
) -> dict[str, Any]:
    row = _resolve(principal, email, "mailboxes:read")
    return service.mailbox_view(row, principal.scopes_for(row.get("email") or ""))


@router.get("/mailboxes/{mailbox_id}", response_model_exclude_none=True)
@audited
async def mailbox_detail(
    mailbox_id: str,
    request: Request,
    principal: Principal = Depends(current_principal),
) -> dict[str, Any]:
    row = _resolve(principal, mailbox_id, "mailboxes:read")
    return service.mailbox_view(row, principal.scopes_for(row.get("email") or ""))


@router.get("/mailboxes/{mailbox_id}/fields")
@audited
async def mailbox_fields(
    mailbox_id: str,
    request: Request,
    principal: Principal = Depends(current_principal),
    profile: str = Query("basic", pattern="^(basic|full|combo)$"),
) -> dict[str, Any]:
    row = _resolve(principal, mailbox_id, "fields:basic")
    return service.fields_view(row, profile, principal.scopes_for(row.get("email") or ""))


# ── 读信 ────────────────────────────────────────────────────────────────────


@router.get("/mailboxes/{mailbox_id}/messages", response_model=MessageListResponse)
@audited
async def mailbox_messages(
    mailbox_id: str,
    request: Request,
    principal: Principal = Depends(current_principal),
    folder: str = Query("inbox", pattern="^(inbox|junkemail)$"),
    limit: int = Query(20, ge=1, le=50),
    after: str = Query("", description="只要这个 ISO 时间之后收到的信"),
    from_: str = Query("", alias="from", description="发件人包含"),
    subject_contains: str = Query(""),
    mode: str = Query("auto", pattern="^(auto|graph|outlook_rest)$"),
) -> dict[str, Any]:
    row = _resolve(principal, mailbox_id, "messages:read")
    service.ensure_not_incubating(row, principal)
    return service.list_messages(
        row, folder=folder, limit=limit, after=after,
        sender=from_, subject_contains=subject_contains, mode=mode,
    )


@router.get("/mailboxes/{mailbox_id}/messages/{message_id}", response_model=MessageItem)
@audited
async def mailbox_message(
    mailbox_id: str,
    message_id: str,
    request: Request,
    principal: Principal = Depends(current_principal),
    mode: str = Query("auto", pattern="^(auto|graph|outlook_rest)$"),
) -> dict[str, Any]:
    row = _resolve(principal, mailbox_id, "messages:read")
    service.ensure_not_incubating(row, principal)
    return service.get_message(row, message_id, mode=mode)


@router.get("/mailboxes/{mailbox_id}/otp", response_model=OtpResponse,
            response_model_exclude_none=True)
@audited
async def mailbox_otp(
    mailbox_id: str,
    request: Request,
    principal: Principal = Depends(current_principal),
    wait_seconds: int = Query(60, ge=0, le=service.MAX_WAIT_SECONDS),
    after: str = Query(""),
    sender: str = Query(""),
    subject_contains: str = Query(""),
    mode: str = Query("auto", pattern="^(auto|graph|outlook_rest)$"),
) -> dict[str, Any]:
    row = _resolve(principal, mailbox_id, "otp:read")
    service.ensure_not_incubating(row, principal)
    return service.wait_otp(
        row, wait_seconds=wait_seconds, after=after,
        sender=sender, subject_contains=subject_contains, mode=mode,
    )


# ── 运维 ────────────────────────────────────────────────────────────────────


@router.get("/health", response_model=HealthResponse)
@audited
async def health(
    request: Request,
    principal: Principal = Depends(current_principal),
) -> dict[str, Any]:
    return {
        "ok": True,
        "readable_count": service.readable_count(principal),
        "schema_version": db.SCHEMA_VERSION,
    }
