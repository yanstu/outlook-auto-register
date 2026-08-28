"""Cloudflare 域名 catch-all 邮箱收码后端（proofs 恢复邮箱的 CF 版）。

与 IMAP 恢复邮箱池的区别：
  - IMAP 池：预置一堆 `email----password` 老号，微软发码到收件箱，用账密登 IMAP 读码。
  - CF 域名（本模块）：不预置账密，按需在自有域名（catch-all）下生成随机地址
    `xxxx@your-cf-domain.com`，微软发码到该地址；邮件经 Cloudflare Email Routing 落到
    自建 CF Worker 临时邮服（dreamhunter2333/cloudflare_temp_email 同款），
    通过其管理 API 按收件地址读出安全验证码。

CF Worker API 形状（本机 account_manager / domain-email-pool / chatgpt-free-hub 三处一致）：
  - 鉴权：请求头 `x-admin-auth: <管理员密码>`
  - 收信：GET  {base}/admin/mails?limit=&offset=&address=<收件地址>
          → list 或 {results|items|data|mails|messages: [...]}
  - 建址（可选）：POST {base}/admin/new_address {enablePrefix,name,domain} → {email}

catch-all 有两种投递拓扑，本模块都兼容：
  A) Email Routing catch-all → 直接进 Worker，按真实收件人存储
     → GET /admin/mails?address=<alias> 直接命中。
  B) catch-all → 转发到单一中转箱（OUTLOOK_CF_FORWARD_EMAIL，如 forward@your-cf-domain.com）
     → GET /admin/mails?address=<forward> 拉全部，再按 alias 过滤 To/Delivered-To 等头。

环境变量（密码等敏感项只从 env 读，绝不硬编码）：
  OUTLOOK_RECOVERY_BACKEND        imap（默认）| cf_domain —— 收码后端开关
  OUTLOOK_CF_DOMAIN               catch-all 域名，默认 your-cf-domain.com
  OUTLOOK_CF_WORKER_API_URL       CF Worker API 根，如 https://apimail.your-cf-domain.com
  OUTLOOK_CF_WORKER_ADMIN_TOKEN   CF Worker 管理员密码（必填）
  OUTLOOK_CF_FORWARD_EMAIL        可选：拓扑 B 的中转箱地址（兜底查询）
  OUTLOOK_CF_PREFIX_LEN           随机前缀长度，默认 12
  OUTLOOK_CF_USE_NEW_ADDRESS      0=随机 catch-all 地址(默认) | 1=调 /admin/new_address 注册
  OUTLOOK_CF_VERIFY_SSL           1=校验证书(默认) | 0=关闭校验
  OUTLOOK_CF_RECOVERY_PLACEHOLDER 六段 combo 第 6 段占位（CF 无账密），默认 cf_domain
"""
from __future__ import annotations

import email.utils
import logging
import os
import random
import re
import string
import time
from dataclasses import dataclass
from email import policy
from email.parser import Parser
from typing import Any, Optional

import requests

# 复用 IMAP 收码同款的微软发件人白名单与验证码抽取，保证解析口径一致
from .mail_reader import _SECURITY_SENDERS, _extract_code

logger = logging.getLogger(__name__)

DEFAULT_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)

# 微软安全码邮件的强特征关键词（发件人解析失败时的兜底匹配）
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


# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------

def recovery_backend() -> str:
    """proofs 恢复邮箱收码后端：``imap``（默认）/ ``cf_domain`` / ``coolhs_mail``。"""
    raw = os.environ.get("OUTLOOK_RECOVERY_BACKEND", "imap").strip().lower()
    if raw in ("cf", "cfdomain", "cf_domain", "cfworker", "catchall", "catch_all"):
        return "cf_domain"
    if raw in ("coolhs", "coolhs_mail", "coolhsmail", "hook_coolhs", "hook.coolhs"):
        return "coolhs_mail"
    return "imap"


def cf_domain_backend_active() -> bool:
    return recovery_backend() == "cf_domain"


@dataclass
class CFConfig:
    api_url: str
    admin_token: str
    domain: str
    forward_email: str = ""
    prefix_len: int = 12
    use_new_address: bool = False
    verify_ssl: bool = True
    alias_limit: int = 50
    forward_limit: int = 100

    def is_valid(self) -> tuple[bool, str]:
        if not self.api_url:
            return False, "OUTLOOK_CF_WORKER_API_URL 未配置"
        if not self.admin_token:
            return False, "OUTLOOK_CF_WORKER_ADMIN_TOKEN 未配置（管理员密码）"
        if not self.domain or "@" in self.domain:
            return False, "OUTLOOK_CF_DOMAIN 未配置或格式错误"
        return True, ""


def _env_first(*names: str, default: str = "") -> str:
    for n in names:
        v = os.environ.get(n, "")
        if v is not None and str(v).strip():
            return str(v).strip()
    return default


def load_config() -> CFConfig:
    try:
        prefix_len = int(os.environ.get("OUTLOOK_CF_PREFIX_LEN", "12"))
    except ValueError:
        prefix_len = 12
    prefix_len = max(4, min(prefix_len, 32))

    def _int_env(name: str, default: int) -> int:
        try:
            return max(1, min(int(os.environ.get(name, str(default))), 500))
        except ValueError:
            return default

    return CFConfig(
        api_url=_env_first("OUTLOOK_CF_WORKER_API_URL", "CFWORKER_API_URL").rstrip("/"),
        admin_token=_env_first("OUTLOOK_CF_WORKER_ADMIN_TOKEN", "CFWORKER_ADMIN_TOKEN"),
        domain=_env_first("OUTLOOK_CF_DOMAIN", "CFWORKER_DOMAIN", "CATCHALL_DOMAIN", default="your-cf-domain.com").lower(),
        forward_email=_env_first("OUTLOOK_CF_FORWARD_EMAIL", "CFWORKER_FORWARD_EMAIL").lower(),
        prefix_len=prefix_len,
        use_new_address=os.environ.get("OUTLOOK_CF_USE_NEW_ADDRESS", "0").strip() == "1",
        verify_ssl=os.environ.get("OUTLOOK_CF_VERIFY_SSL", "1").strip() != "0",
        alias_limit=_int_env("OUTLOOK_CF_MAIL_LIMIT", 50),
        forward_limit=_int_env("OUTLOOK_CF_FORWARD_LIMIT", 100),
    )


def cf_configured() -> bool:
    ok, _ = load_config().is_valid()
    return ok


# ---------------------------------------------------------------------------
# 邮件字段解析（兼容 CF Worker 各种返回形状）
# ---------------------------------------------------------------------------

def _normalize_mail_list(data: Any) -> list[dict]:
    if isinstance(data, list):
        return [m for m in data if isinstance(m, dict)]
    if isinstance(data, dict):
        for key in ("results", "items", "data", "mails", "messages", "value"):
            val = data.get(key)
            if isinstance(val, list):
                return [m for m in val if isinstance(m, dict)]
            if isinstance(val, dict):
                nested = _normalize_mail_list(val)
                if nested:
                    return nested
    return []


def _header_value(raw: str, name: str) -> str:
    m = re.search(rf"(?im)^{re.escape(name)}:\s*(.+)$", raw or "")
    return m.group(1).strip() if m else ""


def _mail_id(mail: dict) -> str:
    for key in ("id", "_id", "message_id", "mail_id"):
        v = mail.get(key)
        if v not in (None, ""):
            return str(v)
    return ""


def _mail_id_sortkey(mail: dict) -> int:
    mid = _mail_id(mail)
    try:
        return int(mid)
    except (TypeError, ValueError):
        return 0


def _mail_subject(mail: dict) -> str:
    for key in ("subject", "title"):
        v = str(mail.get(key) or "").strip()
        if v:
            return v
    return _header_value(str(mail.get("raw") or mail.get("source") or mail.get("mime") or ""), "Subject")


def _mail_from(mail: dict) -> str:
    # CF Worker 列表项用 source 存发件人（account-security-noreply@...）
    for key in ("source", "from", "sender", "from_address", "fromEmail", "mail_from", "mailFrom"):
        v = mail.get(key)
        if isinstance(v, dict):
            v = v.get("address") or v.get("email") or v.get("value") or ""
        v = str(v or "").strip().lower()
        if v:
            return v
    return _header_value(
        str(mail.get("raw") or mail.get("source") or mail.get("mime") or ""), "From"
    ).lower()


def _strip_html(text: str) -> str:
    t = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    t = re.sub(r"<script[\s\S]*?</script>", " ", t, flags=re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"&nbsp;", " ", t, flags=re.I)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def _mail_body(mail: dict) -> tuple[str, str]:
    """返回 (原文, 可见纯文本)。优先结构化字段，回退整封 MIME raw 解析。"""
    for key in ("bodyPreview", "snippet", "text", "text_content", "body", "content", "preview"):
        v = str(mail.get(key) or "").strip()
        if v:
            return v, (_strip_html(v) if "<" in v else v)
    for key in ("html", "html_content", "body_html", "bodyHtml", "raw_html", "content_html"):
        v = str(mail.get(key) or "").strip()
        if v:
            return v, _strip_html(v)

    raw = str(mail.get("raw") or mail.get("mime") or mail.get("message") or mail.get("source") or "")
    if not raw:
        return "", ""
    try:
        msg = Parser(policy=policy.default).parsestr(raw)
        plain: list[str] = []
        html: list[str] = []
        for part in (msg.walk() if msg.is_multipart() else [msg]):
            ctype = (part.get_content_type() or "").lower()
            if ctype not in ("text/plain", "text/html"):
                continue
            try:
                content = str(part.get_content() or "")
            except Exception:  # noqa: BLE001
                payload = part.get_payload(decode=True) or b""
                content = payload.decode(part.get_content_charset() or "utf-8", "ignore") if isinstance(payload, bytes) else str(payload or "")
            content = content.strip()
            if not content:
                continue
            (plain if ctype == "text/plain" else html).append(content)
        if plain:
            body = "\n".join(plain)
            return body, (body if "<" not in body else _strip_html(body))
        if html:
            body = "\n".join(html)
            return body, _strip_html(body)
    except Exception:  # noqa: BLE001
        pass
    sep = "\r\n\r\n" if "\r\n\r\n" in raw else "\n\n"
    idx = raw.find(sep)
    body = raw[idx + len(sep):] if idx >= 0 else raw
    return body, _strip_html(body)


def _parse_cf_timestamp(value: Any) -> float:
    """解析 CF Worker 时间戳。无后缀的 ``YYYY-MM-DD HH:MM:SS`` 按 UTC 处理。

    与 email-code-worker ``icloud-cfworker.ts`` 的 ``normalizeCreatedAt`` 一致：
    CF 返回的 created_at 是 UTC 字符串但不带 ``Z``，若按本机时区解析会在 UTC+8 上差 8 小时，
    导致 since_ts 过滤误跳过刚到的微软验证码（实测随机 alias 地址）。
    """
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        ts = float(value)
        return ts / 1000.0 if ts > 10_000_000_000 else ts
    s = str(value).strip()
    if not s:
        return 0.0
    if s.isdigit():
        ts = float(s)
        return ts / 1000.0 if ts > 10_000_000_000 else ts
    try:
        return email.utils.parsedate_to_datetime(s).timestamp()
    except Exception:  # noqa: BLE001
        pass
    from datetime import datetime, timezone

    iso = s.replace(" ", "T")
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}", s):
        iso = f"{iso}Z"
    elif not iso.endswith("Z") and not re.search(r"[+-]\d{2}:?\d{2}$", iso):
        iso = f"{iso}Z"
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.timestamp()
    except Exception:  # noqa: BLE001
        return 0.0


def _mail_received_ts(mail: dict) -> float:
    for key in ("created_at", "createdAt", "received_at", "receivedDateTime", "date", "timestamp"):
        v = mail.get(key)
        if v in (None, ""):
            continue
        ts = _parse_cf_timestamp(v)
        if ts:
            return ts
    raw = str(mail.get("raw") or mail.get("mime") or "")
    date_hdr = _header_value(raw, "Date")
    if date_hdr:
        ts = _parse_cf_timestamp(date_hdr)
        if ts:
            return ts
    return 0.0


def _mail_recipient_blob(mail: dict) -> str:
    """把邮件里所有能表示收件地址的字段拼成一串，用于 alias 匹配判定。"""
    parts: list[str] = []
    for key in (
        "original_recipient", "originalRecipient", "envelope_to", "envelopeTo",
        "to", "to_address", "toAddress", "mail_to", "mailTo", "recipient",
        "receiver", "receivers", "address",
    ):
        v = mail.get(key)
        if isinstance(v, dict):
            v = v.get("address") or v.get("email") or v.get("value") or ""
        if isinstance(v, (list, tuple)):
            for item in v:
                if isinstance(item, dict):
                    item = item.get("address") or item.get("email") or item.get("value") or ""
                if item:
                    parts.append(str(item))
            continue
        if v:
            parts.append(str(v))
    raw = str(mail.get("raw") or mail.get("source") or mail.get("mime") or "")
    for header in ("To", "Delivered-To", "X-Original-To", "X-Forwarded-To", "Envelope-To", "Cc"):
        hv = _header_value(raw, header)
        if hv:
            parts.append(hv)
    m = re.search(r"(?i)\bfor\s+<([^>]+)>", raw)
    if m:
        parts.append(m.group(1))
    return " ".join(parts).lower()


def _alias_matches(mail: dict, alias: str) -> bool:
    target = alias.strip().lower()
    if not target:
        return False
    blob = _mail_recipient_blob(mail)
    if target in blob:
        return True
    raw = str(mail.get("raw") or mail.get("source") or mail.get("mime") or "").lower()
    return target in raw


def _mail_has_recipient_info(mail: dict) -> bool:
    return bool(_mail_recipient_blob(mail).strip())


def _is_ms_security_mail(mail: dict) -> bool:
    frm = _mail_from(mail)
    if any(s in frm for s in _SECURITY_SENDERS):
        return True
    subject = _mail_subject(mail).lower()
    _, visible = _mail_body(mail)
    blob = f"{subject}\n{visible}".lower()
    return any(k in blob for k in _MS_KEYWORDS)


# ---------------------------------------------------------------------------
# CF Worker 客户端
# ---------------------------------------------------------------------------

class CFDomainMailClient:
    """自建 CF Worker 临时邮服（catch-all）读取客户端。"""

    def __init__(self, cfg: Optional[CFConfig] = None, *, timeout: int = 20):
        self.cfg = cfg or load_config()
        ok, err = self.cfg.is_valid()
        if not ok:
            raise ValueError(f"CF 域名邮箱配置无效：{err}")
        self.api = self.cfg.api_url.rstrip("/")
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
            "x-admin-auth": self.cfg.admin_token,
        }

    def list_mails(self, address: str, *, limit: int = 50, offset: int = 0) -> list[dict]:
        try:
            r = self.session.get(
                f"{self.api}/admin/mails",
                params={"limit": limit, "offset": offset, "address": address},
                headers=self._headers(),
                timeout=self.timeout,
            )
        except Exception as exc:  # noqa: BLE001
            self.last_error = f"{type(exc).__name__}: {exc}"
            self.last_count = 0
            return []
        if r.status_code >= 400:
            body = (r.text or "")[:200]
            if r.status_code == 401 or "admin password" in body.lower():
                self.last_error = "CF Worker 管理员密码错误（x-admin-auth 401）"
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
        mails = _normalize_mail_list(data)
        self.last_error = ""
        self.last_count = len(mails)
        return mails

    def create_address(self, *, name: str = "", enable_prefix: bool = True) -> str:
        local = name or _random_local(self.cfg.prefix_len)
        payload: dict[str, Any] = {"enablePrefix": enable_prefix, "name": local, "domain": self.cfg.domain}
        r = self.session.post(
            f"{self.api}/admin/new_address",
            json=payload,
            headers=self._headers(),
            timeout=self.timeout,
        )
        if r.status_code >= 400:
            raise RuntimeError(f"CF Worker 建址失败 HTTP {r.status_code}: {(r.text or '')[:200]}")
        try:
            data = r.json()
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"CF Worker 建址响应非 JSON: {(r.text or '')[:120]}") from exc
        email_addr = str((data or {}).get("email") or (data or {}).get("address") or "").strip()
        if not email_addr:
            raise RuntimeError(f"CF Worker 建址未返回地址: {data!r}")
        return email_addr

    def health_check(self) -> tuple[bool, str]:
        probe = self.cfg.forward_email or f"__healthcheck__@{self.cfg.domain}"
        mails = self.list_mails(probe, limit=1)
        if self.last_error:
            return False, self.last_error
        return True, f"API 可达，探测地址 {probe} 返回 {len(mails)} 条"

    # ── 收码 ──

    def _query_addresses(self, alias: str) -> list[tuple[str, bool, int]]:
        """返回 [(查询地址, strict, limit), ...]。strict=True 表示必须严格匹配 alias。

        直查 alias 的专用地址通常很空，limit 小即可；中转箱（forward）是多别名共用的
        繁忙收件箱（实测上千封），需要更大的 limit 才能兜住最新那封微软验证码。
        """
        out: list[tuple[str, bool, int]] = [(alias, False, self.cfg.alias_limit)]
        fwd = (self.cfg.forward_email or "").strip().lower()
        if fwd and fwd != alias.strip().lower():
            out.append((fwd, True, self.cfg.forward_limit))
        return out

    def _candidate_mails(self, alias: str) -> list[dict]:
        seen: set[str] = set()
        merged: list[dict] = []
        # new_address 注册的邮箱：与 CF 管理台一致，直查 address=alias（截图里「查询」按钮）
        queries: list[tuple[str, bool, int]] = [(alias, False, self.cfg.alias_limit)]
        if not self.cfg.use_new_address:
            fwd = (self.cfg.forward_email or "").strip().lower()
            if fwd and fwd != alias.strip().lower():
                queries.append((fwd, True, self.cfg.forward_limit))
        for query_addr, strict, limit in queries:
            mails = self.list_mails(query_addr, limit=limit)
            for mail in mails:
                if strict:
                    if not _alias_matches(mail, alias):
                        continue
                else:
                    if _mail_has_recipient_info(mail) and not _alias_matches(mail, alias):
                        continue
                mid = _mail_id(mail) or f"_noid_{id(mail)}"
                if mid in seen:
                    continue
                seen.add(mid)
                merged.append(mail)
        merged.sort(key=_mail_id_sortkey, reverse=True)
        return merged

    def snapshot_ids(self, alias: str) -> set[str]:
        """快照该 alias 专用地址已有邮件 id（与 CF 管理台直查 address=alias 一致）。"""
        mails = self.list_mails(alias, limit=self.cfg.alias_limit)
        return {mid for mail in mails if (mid := _mail_id(mail))}

    def read_security_code(
        self,
        alias: str,
        *,
        since_ts: float = 0.0,
        before_ids: Optional[set[str]] = None,
        timeout: int = 150,
        poll_interval: float = 4.0,
    ) -> str:
        """轮询 CF Worker，读该 alias 收到的最新一封微软安全验证码。

        与 email-code 的 icloud_cfworker / CF 管理台一致：new_address 邮箱直查
        ``GET /admin/mails?address=<alias>``；仅旧 catch-all 才查中转箱并按 To 头过滤。

        注意：只有成功提取验证码后才把 id 记入 consumed，避免首 poll 正文未就绪时
        误标记已读导致后续 150s 全跳过（实测 alias@your-cf-domain.com 信已到但 pipeline 超时）。
        """
        skip_ids: set[str] = set(before_ids or [])
        consumed: set[str] = set()
        deadline = time.time() + timeout
        skipped_snapshot = 0
        skipped_non_ms = 0
        while time.time() < deadline:
            for mail in self._candidate_mails(alias):
                mid = _mail_id(mail)
                if mid and mid in skip_ids:
                    skipped_snapshot += 1
                    continue
                if mid and mid in consumed:
                    continue
                rcv = _mail_received_ts(mail)
                # since_ts 仅过滤「快照前」的旧信；新信 id 不在 before_ids 时不应被误杀
                if since_ts and rcv and rcv < since_ts - 120 and mid and mid in skip_ids:
                    continue
                if not _is_ms_security_mail(mail):
                    skipped_non_ms += 1
                    continue
                subject = _mail_subject(mail)
                _, visible = _mail_body(mail)
                code = _extract_code(subject, visible)
                if code:
                    if mid:
                        consumed.add(mid)
                    logger.info("CF 恢复邮箱 %s 读到微软 OTT=%s (mail_id=%s)", alias, code, mid or "?")
                    return code
            time.sleep(poll_interval)
        detail = self.last_error or (
            f"直查 {alias} 最近 {self.last_count} 封未提取到验证码"
            + (f"；快照跳过 {skipped_snapshot} 封" if skipped_snapshot else "")
            + (f"；非微软信 {skipped_non_ms} 封" if skipped_non_ms else "")
        )
        logger.error("CF 恢复邮箱 %s 等待 OTT 超时(%ss)：%s", alias, timeout, detail)
        return ""


# ---------------------------------------------------------------------------
# 地址分配
# ---------------------------------------------------------------------------

def _random_local(length: int) -> str:
    first = random.choice(string.ascii_lowercase)
    rest = "".join(random.choice(string.ascii_lowercase + string.digits) for _ in range(max(1, length - 1)))
    return first + rest


def allocate_address(client: Optional[CFDomainMailClient] = None) -> str:
    """按需分配一个 catch-all 恢复邮箱地址。

    默认直接生成随机 ``xxxx@域名``（catch-all 无需注册）；
    若 OUTLOOK_CF_USE_NEW_ADDRESS=1 则调 CF Worker /admin/new_address 注册地址。
    """
    cfg = client.cfg if client else load_config()
    ok, err = cfg.is_valid()
    if not ok:
        raise ValueError(f"CF 域名邮箱配置无效：{err}")
    if cfg.use_new_address:
        c = client or CFDomainMailClient(cfg)
        return c.create_address()
    return f"{_random_local(cfg.prefix_len)}@{cfg.domain}"


def recovery_placeholder() -> str:
    """六段 combo 第 6 段占位（CF 恢复邮箱无账密，用非敏感标记以保留六段导出）。"""
    return os.environ.get("OUTLOOK_CF_RECOVERY_PLACEHOLDER", "cf_domain").strip() or "cf_domain"
