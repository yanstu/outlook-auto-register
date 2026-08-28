"""Mailbox API 的统一错误体：{code, message, ...}，不套 FastAPI 默认的 detail 外壳。"""
from __future__ import annotations

from typing import Any


class MailboxApiError(Exception):
    """带 HTTP 状态码与机器可读 code 的业务错误。"""

    def __init__(self, status_code: int, code: str, message: str, **extra: Any) -> None:
        super().__init__(message)
        self.status_code = int(status_code)
        self.code = code
        self.message = message
        self.extra: dict[str, Any] = {k: v for k, v in extra.items() if v is not None}

    def body(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"code": self.code, "message": self.message}
        payload.update(self.extra)
        return payload


def unauthorized(message: str = "缺少或无效的访问令牌", code: str = "unauthorized") -> MailboxApiError:
    return MailboxApiError(401, code, message)


def forbidden(message: str, code: str = "forbidden", **extra: Any) -> MailboxApiError:
    return MailboxApiError(403, code, message, **extra)


def not_found(message: str, code: str = "not_found", **extra: Any) -> MailboxApiError:
    return MailboxApiError(404, code, message, **extra)
