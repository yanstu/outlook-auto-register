"""账号生命周期：孵化期、长存筛选。

新号默认进入孵化期（OUTLOOK_INCUBATION_HOURS，默认 48h），期间批量测活 /
保活脚本应跳过，避免对新号高频打微软接口。
"""
from __future__ import annotations

import os
from datetime import datetime, timedelta
from typing import Any, Optional

INCUBATING_TAG = "incubating"
DEFAULT_INCUBATION_HOURS = 48


def incubation_hours() -> float:
    raw = (os.environ.get("OUTLOOK_INCUBATION_HOURS") or "").strip()
    if not raw:
        return float(DEFAULT_INCUBATION_HOURS)
    try:
        hours = float(raw)
    except ValueError:
        return float(DEFAULT_INCUBATION_HOURS)
    return max(0.0, hours)


def parse_created_at(value: Any) -> Optional[datetime]:
    """解析账号 created_at（ISO 或常见变体）。失败返回 None。"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None) if value.tzinfo else value
    text = str(value).strip()
    if not text:
        return None
    # 允许尾部 Z / 带空格
    text = text.replace("Z", "").replace(" ", "T", 1) if " " in text and "T" not in text else text.replace("Z", "")
    for candidate in (text, text[:19], text.split("+")[0].split(".")[0]):
        try:
            return datetime.fromisoformat(candidate)
        except ValueError:
            continue
    return None


def incubation_until(created_at: Any, *, hours: Optional[float] = None) -> Optional[datetime]:
    """返回孵化结束时刻；无 created_at 则无法计算。"""
    created = parse_created_at(created_at)
    if created is None:
        return None
    h = incubation_hours() if hours is None else max(0.0, float(hours))
    return created + timedelta(hours=h)


def is_incubating(
    created_at: Any,
    *,
    now: Optional[datetime] = None,
    hours: Optional[float] = None,
) -> bool:
    """当前是否仍在孵化期内。"""
    until = incubation_until(created_at, hours=hours)
    if until is None:
        return False
    current = now or datetime.now()
    if current.tzinfo:
        current = current.replace(tzinfo=None)
    return current < until


def account_age_days(
    created_at: Any,
    *,
    now: Optional[datetime] = None,
) -> Optional[float]:
    created = parse_created_at(created_at)
    if created is None:
        return None
    current = now or datetime.now()
    if current.tzinfo:
        current = current.replace(tzinfo=None)
    return max(0.0, (current - created).total_seconds() / 86400.0)


def enrich_lifecycle_fields(
    row: dict[str, Any],
    *,
    now: Optional[datetime] = None,
    hours: Optional[float] = None,
) -> dict[str, Any]:
    """就地写入 incubating / incubation_until，并同步 tags 中的 incubating 标记。"""
    created = row.get("created_at")
    until = incubation_until(created, hours=hours)
    incubating = is_incubating(created, now=now, hours=hours)
    row["incubating"] = incubating
    row["incubation_until"] = until.isoformat(timespec="seconds") if until else ""

    tags = list(row.get("tags") or [])
    has_tag = INCUBATING_TAG in tags
    if incubating and not has_tag:
        tags.append(INCUBATING_TAG)
    elif not incubating and has_tag:
        tags = [t for t in tags if t != INCUBATING_TAG]
    row["tags"] = tags
    return row


def is_long_lived(
    row: dict[str, Any],
    *,
    min_days: float = 7.0,
    now: Optional[datetime] = None,
) -> bool:
    """满足长存门槛：非孵化、有 recovery combo、年龄 ≥ min_days。"""
    if is_incubating(row.get("created_at"), now=now):
        return False
    age = account_age_days(row.get("created_at"), now=now)
    if age is None or age < float(min_days):
        return False
    combo = (row.get("combo_recovery") or "").strip()
    if combo and len(combo.split("----")) >= 6:
        return True
    # 可从字段拼装
    if row.get("recovery_email") and row.get("refresh_token"):
        return True
    return False


def combo_recovery_line(row: dict[str, Any]) -> str:
    """导出用六段 recovery 行；不可用则空串。"""
    combo = (row.get("combo_recovery") or "").strip()
    if combo and len(combo.split("----")) >= 6:
        return combo
    email = (row.get("email") or "").strip()
    pwd = row.get("password") or ""
    cid = row.get("client_id") or ""
    rt = row.get("refresh_token") or ""
    rec_e = row.get("recovery_email") or ""
    rec_p = row.get("recovery_password") or ""
    if email and rt and rec_e:
        return "----".join([email, pwd, cid, rt, rec_e, rec_p])
    return ""
