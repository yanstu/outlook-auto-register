"""外部恢复邮箱池（login.exe 同款：your-recovery-host.com 等第三方收码邮箱）。

每行：`recovery_email----recovery_password`
环境变量：
  OUTLOOK_EXTERNAL_RECOVERY_POOL_FILE — 池文件路径
  OUTLOOK_RECOVERY_IMAP_HOST — IMAP 主机（必填，如 imap.your-recovery-host.com）
  OUTLOOK_RECOVERY_IMAP_PORT — 默认 993

收码后端可切换（OUTLOOK_RECOVERY_BACKEND）：
  imap（默认） — 本模块的第三方 IMAP 恢复邮箱池。
  cf_domain    — Cloudflare 域名 catch-all 邮箱（见 cf_domain_mail.py），无需预置账密，
                 按需在自有域名下生成随机地址并经 CF Worker API 收码。
  coolhs_mail  — 自建 coolhs-mail（hook.coolhs.com，见 coolhs_mail.py），
                 ``x-api-token`` + ``/api/mailbox/...``，域名如 mail.coolhs.com。
"""
from __future__ import annotations

import logging
import os
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional

logger = logging.getLogger(__name__)

_lock = threading.Lock()
_cursor = 0


@dataclass
class ExternalRecovery:
    email: str
    password: str

    def masked(self) -> str:
        name, _, dom = self.email.partition("@")
        head = name[:2] if len(name) > 2 else name[:1]
        return f"{head}***@{dom}"


def pool_path() -> Optional[Path]:
    env = os.environ.get("OUTLOOK_EXTERNAL_RECOVERY_POOL_FILE", "").strip()
    if not env:
        return None
    p = Path(env).expanduser()
    return p if p.exists() else None


def imap_host() -> str:
    return os.environ.get("OUTLOOK_RECOVERY_IMAP_HOST", "").strip()


def imap_port() -> int:
    try:
        return int(os.environ.get("OUTLOOK_RECOVERY_IMAP_PORT", "993"))
    except ValueError:
        return 993


def load_pool() -> list[ExternalRecovery]:
    p = pool_path()
    if not p:
        return []
    out: list[ExternalRecovery] = []
    for line in p.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("----")
        if len(parts) < 2:
            continue
        email, pwd = parts[0].strip(), parts[1].strip()
        if "@" in email and pwd:
            out.append(ExternalRecovery(email, pwd))
    logger.info("外部恢复邮箱池载入 %d 个（来源 %s）", len(out), p)
    return out


def iter_accounts(limit: int = 8) -> Iterator[ExternalRecovery]:
    global _cursor
    pool = load_pool()
    if not pool:
        return
    with _lock:
        start = _cursor % len(pool)
        _cursor = (_cursor + 1) % len(pool)
    n = min(limit, len(pool))
    for i in range(n):
        yield pool[(start + i) % len(pool)]


def recovery_backend() -> str:
    """proofs 恢复邮箱收码后端：``imap``（默认）/ ``cf_domain`` / ``coolhs_mail``。"""
    from . import cf_domain_mail
    return cf_domain_mail.recovery_backend()


def external_pool_enabled() -> bool:
    """proofs 恢复邮箱功能是否可用（供 webapp / ss_post 判定是否能绑定恢复邮箱）。

    - imap 后端：需 OUTLOOK_EXTERNAL_RECOVERY_POOL_FILE + OUTLOOK_RECOVERY_IMAP_HOST。
    - cf_domain 后端：需 CF Worker API/域名/管理员密码齐全（见 cf_domain_mail.cf_configured）。
    - coolhs_mail 后端：需 COOLHS_MAIL_BASE_URL + COOLHS_MAIL_API_TOKEN + COOLHS_MAIL_DOMAIN。
    """
    if os.environ.get("OUTLOOK_EXTERNAL_RECOVERY", "1") == "0":
        return False
    backend = recovery_backend()
    if backend == "cf_domain":
        from . import cf_domain_mail
        return cf_domain_mail.cf_configured()
    if backend == "coolhs_mail":
        from . import coolhs_mail
        return coolhs_mail.coolhs_configured()
    return bool(pool_path() and imap_host())
