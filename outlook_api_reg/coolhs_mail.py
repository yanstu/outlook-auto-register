"""coolhs-mail（hook.coolhs.com）恢复邮箱收码后端。

与 ``cf_domain_mail``（CF Worker ``/admin/mails``）并列：本仓库基础设施用的是
coolhs-mail（``x-api-token`` + ``/api/mailbox/...``），见 cheliz ``deploy/coolhs-mail/USAGE.md``。

环境变量：
  OUTLOOK_RECOVERY_BACKEND=coolhs_mail
  COOLHS_MAIL_BASE_URL=https://hook.coolhs.com
  COOLHS_MAIL_API_TOKEN=...
  COOLHS_MAIL_DOMAIN=mail.coolhs.com
  COOLHS_MAIL_PREFIX_LEN=12          （可选）
  COOLHS_MAIL_RECOVERY_PLACEHOLDER=coolhs_mail  （六段 combo 第 6 段占位）
  COOLHS_MAIL_VERIFY_SSL=1
"""
from __future__ import annotations

import logging
import os
import random
import string
import time
from dataclasses import dataclass
from typing import Any, Optional
from urllib.parse import quote

import requests

from .mail_reader import _SECURITY_SENDERS, _extract_code

logger = logging.getLogger(__name__)

DEFAULT_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)

_MS_KEYWORDS = (
    "security code",
    "verification code",
    "single-use code",
    "verify your email",
    "microsoft account",
    "your code is",
    "安全代码",
    "验证码",
    "安全码",
)


@dataclass
class CoolhsConfig:
    base_url: str
    api_token: str
    domain: str
    prefix_len: int = 12
    verify_ssl: bool = True
    use_create_api: bool = True

    def is_valid(self) -> tuple[bool, str]:
        if not self.base_url:
            return False, "COOLHS_MAIL_BASE_URL 未配置"
        if not self.api_token:
            return False, "COOLHS_MAIL_API_TOKEN 未配置"
        if not self.domain or "@" in self.domain:
            return False, "COOLHS_MAIL_DOMAIN 未配置或格式错误"
        return True, ""


def _env_first(*names: str, default: str = "") -> str:
    for n in names:
        v = os.environ.get(n, "")
        if v is not None and str(v).strip():
            return str(v).strip()
    return default


def load_config() -> CoolhsConfig:
    try:
        prefix_len = int(os.environ.get("COOLHS_MAIL_PREFIX_LEN", "12"))
    except ValueError:
        prefix_len = 12
    prefix_len = max(4, min(prefix_len, 32))
    return CoolhsConfig(
        base_url=_env_first(
            "COOLHS_MAIL_BASE_URL", "OUTLOOK_COOLHS_MAIL_BASE_URL",
            default="https://hook.coolhs.com",
        ).rstrip("/"),
        api_token=_env_first("COOLHS_MAIL_API_TOKEN", "OUTLOOK_COOLHS_MAIL_API_TOKEN"),
        domain=_env_first(
            "COOLHS_MAIL_DOMAIN", "OUTLOOK_COOLHS_MAIL_DOMAIN",
            default="mail.coolhs.com",
        ).lower().lstrip("@"),
        prefix_len=prefix_len,
        verify_ssl=os.environ.get("COOLHS_MAIL_VERIFY_SSL", "1").strip() != "0",
        use_create_api=os.environ.get("COOLHS_MAIL_USE_CREATE_API", "1").strip() != "0",
    )


def coolhs_configured() -> bool:
    ok, _ = load_config().is_valid()
    return ok


def coolhs_backend_active() -> bool:
    from . import cf_domain_mail
    return cf_domain_mail.recovery_backend() == "coolhs_mail"


def recovery_placeholder() -> str:
    return (
        os.environ.get("COOLHS_MAIL_RECOVERY_PLACEHOLDER", "coolhs_mail").strip()
        or "coolhs_mail"
    )


def _random_local(length: int) -> str:
    first = random.choice(string.ascii_lowercase)
    rest = "".join(
        random.choice(string.ascii_lowercase + string.digits)
        for _ in range(max(1, length - 1))
    )
    return first + rest


class CoolhsMailClient:
    """hook.coolhs.com coolhs-mail API 客户端（proofs 恢复邮箱）。"""

    def __init__(self, cfg: Optional[CoolhsConfig] = None, *, timeout: int = 25):
        self.cfg = cfg or load_config()
        ok, err = self.cfg.is_valid()
        if not ok:
            raise ValueError(f"coolhs-mail 配置无效：{err}")
        self.api = self.cfg.base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": DEFAULT_UA})
        self.session.verify = self.cfg.verify_ssl
        self.last_error = ""
        self.last_count = 0

    def _headers(self) -> dict[str, str]:
        return {
            "accept": "application/json, text/plain, */*",
            "content-type": "application/json",
            "x-api-token": self.cfg.api_token,
        }

    def health_check(self) -> tuple[bool, str]:
        try:
            r = self.session.get(
                f"{self.api}/api/health",
                headers=self._headers(),
                timeout=self.timeout,
            )
        except Exception as exc:  # noqa: BLE001
            return False, f"{type(exc).__name__}: {exc}"
        if r.status_code == 401:
            return False, "coolhs-mail API token 无效（x-api-token 401）"
        if r.status_code >= 400:
            return False, f"HTTP {r.status_code}: {(r.text or '')[:160]}"
        try:
            data = r.json() if isinstance(r.json(), dict) else {}
        except Exception:  # noqa: BLE001
            data = {}
        if not data.get("ok"):
            return False, f"health 非 ok: {(r.text or '')[:160]}"
        domain = data.get("domain") or self.cfg.domain
        return True, f"API 可达 domain={domain} messages={data.get('messages', '?')}"

    def create_address(self, *, prefix: str = "") -> str:
        local_prefix = (prefix or "rh").strip().lower() or "rh"
        try:
            r = self.session.post(
                f"{self.api}/api/mailbox/create",
                headers=self._headers(),
                json={
                    "prefix": local_prefix,
                    "domain": self.cfg.domain,
                    "note": "outlook-auto-register recovery",
                },
                timeout=self.timeout,
            )
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"coolhs-mail 建址网络失败: {exc}") from exc
        if r.status_code == 401:
            raise RuntimeError("coolhs-mail API token 无效（x-api-token 401）")
        if r.status_code >= 400:
            raise RuntimeError(f"coolhs-mail 建址失败 HTTP {r.status_code}: {(r.text or '')[:200]}")
        try:
            data = r.json()
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"coolhs-mail 建址响应非 JSON: {(r.text or '')[:120]}") from exc
        addr = str((data or {}).get("address") or "").strip().lower()
        if not addr or "@" not in addr:
            raise RuntimeError(f"coolhs-mail 建址未返回地址: {data!r}")
        return addr

    def list_messages(self, address: str, *, limit: int = 20, since_id: int = 0) -> list[dict]:
        enc = quote(address, safe="@._-+")
        params: dict[str, Any] = {"limit": max(1, min(limit, 200)), "include_html": "true"}
        if since_id:
            params["since_id"] = since_id
        try:
            r = self.session.get(
                f"{self.api}/api/mailbox/{enc}/messages",
                headers=self._headers(),
                params=params,
                timeout=self.timeout,
            )
        except Exception as exc:  # noqa: BLE001
            self.last_error = f"{type(exc).__name__}: {exc}"
            self.last_count = 0
            return []
        if r.status_code >= 400:
            body = (r.text or "")[:200]
            if r.status_code == 401:
                self.last_error = "coolhs-mail API token 无效（x-api-token 401）"
            else:
                self.last_error = f"HTTP {r.status_code}: {body}"
            self.last_count = 0
            return []
        try:
            data = r.json()
        except Exception:  # noqa: BLE001
            self.last_error = f"响应非 JSON: HTTP {r.status_code} {(r.text or '')[:120]}"
            self.last_count = 0
            return []
        msgs = data.get("messages") if isinstance(data, dict) else None
        if not isinstance(msgs, list):
            msgs = []
        out = [m for m in msgs if isinstance(m, dict)]
        self.last_error = ""
        self.last_count = len(out)
        return out

    def snapshot_ids(self, alias: str) -> set[str]:
        mails = self.list_messages(alias, limit=50)
        return {str(m.get("id")) for m in mails if m.get("id") not in (None, "")}

    def _poll_code_endpoint(self, address: str, *, wait: int, since_id: int = 0) -> Optional[str]:
        enc = quote(address, safe="@._-+")
        params: dict[str, Any] = {"wait": max(0, min(int(wait), 60))}
        if since_id:
            params["since_id"] = since_id
        try:
            r = self.session.get(
                f"{self.api}/api/mailbox/{enc}/code",
                headers=self._headers(),
                params=params,
                timeout=max(self.timeout, int(wait) + 10),
            )
        except Exception as exc:  # noqa: BLE001
            self.last_error = f"{type(exc).__name__}: {exc}"
            return None
        if r.status_code >= 400:
            self.last_error = f"HTTP {r.status_code}: {(r.text or '')[:160]}"
            return None
        try:
            data = r.json()
        except Exception:  # noqa: BLE001
            self.last_error = f"code 响应非 JSON: {(r.text or '')[:120]}"
            return None
        if isinstance(data, dict) and data.get("ok") and data.get("code"):
            return str(data["code"]).strip()
        return None

    @staticmethod
    def _is_ms_security_mail(mail: dict) -> bool:
        frm = str(mail.get("from") or "").lower()
        if any(s in frm for s in _SECURITY_SENDERS):
            return True
        subject = str(mail.get("subject") or "").lower()
        text = str(mail.get("text") or "")
        html = str(mail.get("html") or "")
        blob = f"{subject}\n{text}\n{html}".lower()
        return any(k in blob for k in _MS_KEYWORDS)

    def read_security_code(
        self,
        alias: str,
        *,
        since_ts: float = 0.0,  # noqa: ARG002 — 接口对齐 cf_domain_mail
        before_ids: Optional[set[str]] = None,
        timeout: int = 150,
        poll_interval: float = 4.0,
    ) -> str:
        """长轮询 coolhs-mail，读该 alias 收到的微软安全验证码。"""
        skip_ids: set[str] = set(before_ids or [])
        since_id = 0
        for sid in skip_ids:
            try:
                since_id = max(since_id, int(sid))
            except (TypeError, ValueError):
                pass
        deadline = time.time() + timeout
        while time.time() < deadline:
            remain = max(1, int(deadline - time.time()))
            wait = min(60, remain)
            code = self._poll_code_endpoint(alias, wait=wait, since_id=since_id)
            if code:
                # /code 可能命中任意验证码；再用 messages 确认是否微软安全信
                msgs = self.list_messages(alias, limit=20, since_id=since_id)
                for mail in msgs:
                    mid = str(mail.get("id") or "")
                    if mid and mid in skip_ids:
                        continue
                    if not self._is_ms_security_mail(mail):
                        continue
                    subject = str(mail.get("subject") or "")
                    body = str(mail.get("text") or mail.get("html") or "")
                    extracted = _extract_code(subject, body) or (
                        str(mail.get("code") or "").strip()
                    )
                    if extracted:
                        logger.info(
                            "coolhs-mail %s 读到微软 OTT=%s (mail_id=%s)",
                            alias, extracted, mid or "?",
                        )
                        return extracted
                # 无明确微软信时仍返回 /code 结果（多数场景就是微软码）
                logger.info("coolhs-mail %s /code 命中 OTT=%s", alias, code)
                return code

            # 兜底：短拉 messages + 本库抽取
            msgs = self.list_messages(alias, limit=20, since_id=since_id)
            for mail in msgs:
                mid = str(mail.get("id") or "")
                if mid and mid in skip_ids:
                    continue
                if not self._is_ms_security_mail(mail):
                    continue
                subject = str(mail.get("subject") or "")
                body = str(mail.get("text") or mail.get("html") or "")
                extracted = _extract_code(subject, body) or str(mail.get("code") or "").strip()
                if extracted:
                    logger.info(
                        "coolhs-mail %s messages 读到微软 OTT=%s (mail_id=%s)",
                        alias, extracted, mid or "?",
                    )
                    return extracted
            time.sleep(min(poll_interval, max(0.5, deadline - time.time())))

        detail = self.last_error or f"coolhs-mail {alias} 等待 OTT 超时({timeout}s)"
        logger.error("coolhs-mail %s 等待 OTT 超时：%s", alias, detail)
        return ""


def allocate_address(client: Optional[CoolhsMailClient] = None) -> str:
    """分配 ``xxxx@mail.coolhs.com``（默认调 create API；也可纯 catch-all 随机）。"""
    cfg = client.cfg if client else load_config()
    ok, err = cfg.is_valid()
    if not ok:
        raise ValueError(f"coolhs-mail 配置无效：{err}")
    if cfg.use_create_api:
        c = client or CoolhsMailClient(cfg)
        return c.create_address(prefix="out")
    return f"{_random_local(cfg.prefix_len)}@{cfg.domain}"
