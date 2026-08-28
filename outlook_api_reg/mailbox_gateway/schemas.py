"""Mailbox API 的请求 / 响应模型。"""
from __future__ import annotations

from typing import Any, Optional

from pydantic import BaseModel, Field


class LoginRequest(BaseModel):
    email: str
    password: str


class LoginResponse(BaseModel):
    token: str
    expires_in: int
    expires_at: str
    mailbox_id: str
    email: str
    scopes: list[str]


class CreateKeyRequest(BaseModel):
    name: str = ""
    scopes: Any = Field(default_factory=list, description="逗号分隔字符串或字符串数组")
    grants: Optional[list[str]] = Field(default=None, description="留空 = 授权全部邮箱")
    expires_at: str = ""


class CreateKeyResponse(BaseModel):
    id: str
    key: str
    name: str
    scopes: list[str]
    grants: list[str]


class PrincipalResponse(BaseModel):
    id: str
    kind: str
    name: str
    scopes: list[str]
    all_mailboxes: bool
    mailboxes: list[str]
    session_email: str
    expires_at: str


class MailboxItem(BaseModel):
    mailbox_id: str
    email: str
    batch: Optional[str] = None
    batch_no: Optional[int] = None
    created_at: Optional[str] = None
    incubating: Optional[bool] = None
    incubation_until: Optional[str] = None
    readable: Optional[bool] = None
    has_token: Optional[bool] = None
    last_alive_at: Optional[str] = None
    tags: Optional[list[str]] = None
    password: Optional[str] = None
    client_id: Optional[str] = None
    recovery_email: Optional[str] = None
    recovery_password: Optional[str] = None


class MailboxListResponse(BaseModel):
    total: int
    limit: int
    offset: int
    items: list[MailboxItem]


class MessageItem(BaseModel):
    id: str
    folder: str
    subject: str
    from_: str = Field(default="", alias="from")
    received: str
    preview: str
    body: str

    model_config = {"populate_by_name": True}


class MessageListResponse(BaseModel):
    email: str
    folder: str
    mode: str
    count: int
    messages: list[MessageItem]


class OtpResponse(BaseModel):
    email: str
    found: bool
    code: str
    mode: str
    waited_seconds: Optional[int] = None
    message: Optional[MessageItem] = None


class HealthResponse(BaseModel):
    ok: bool
    readable_count: int
    schema_version: int
