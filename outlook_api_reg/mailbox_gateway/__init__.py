"""Mailbox API v1：给其他注册项目（以及未来的用户端）用 HTTP 收信、取码、查字段。

对外只暴露三样东西：
- `mailbox_router`：挂到 FastAPI 上的 `/api/v1` 路由；
- `install_error_handlers(app)`：把业务错误渲染成 `{code, message, ...}`；
- `api_enabled()` / `bootstrap()`：开关与首次启动的 admin key 写入。
"""
from __future__ import annotations

import logging
import os

from fastapi import FastAPI
from fastapi.responses import JSONResponse

from .errors import MailboxApiError
from .routes import router as mailbox_router
from .store import bootstrap_admin_from_env

logger = logging.getLogger(__name__)

__all__ = [
    "MailboxApiError",
    "api_enabled",
    "bootstrap",
    "install_error_handlers",
    "mailbox_router",
]


def api_enabled() -> bool:
    """OUTLOOK_MAILBOX_API_ENABLED：缺省开启，显式 0 / false / off 关闭。"""
    raw = (os.environ.get("OUTLOOK_MAILBOX_API_ENABLED") or "1").strip().lower()
    return raw not in ("0", "false", "off", "no")


def install_error_handlers(app: FastAPI) -> None:
    async def _handler(_request, exc: MailboxApiError) -> JSONResponse:
        return JSONResponse(status_code=exc.status_code, content=exc.body())

    app.add_exception_handler(MailboxApiError, _handler)


def bootstrap() -> None:
    """启动时把 env 里的 admin key 写进库（幂等；env 为空则什么都不做）。"""
    try:
        if bootstrap_admin_from_env():
            logger.info("Mailbox API：已写入 env 提供的 admin key")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Mailbox API admin key 写入失败: %s", exc)
