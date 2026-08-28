"""外部 Outlook 账号导入：把非本机注册产出的 combo 并入本地账号库。

## 背景

同一台 38 服务器上还跑着另一套系统（qoderji），它的 `email_inventory` 表里攒了几千个
早先已经注册好的 Outlook 邮箱（`email----password----client_id----refresh_token`
四段 combo，`password` 段常是占位符 `x`，真正能用的是 OAuth `refresh_token`）。
这些号是 qoderji 拿来做产品注册用的原料，注册完（`status=consumed`）或还没用
（`status=untried`）都仍然是能收信的正常 Outlook 邮箱；只有 `status=dead`
（OAuth 永久失效，如 `invalid_grant`）才是真正报废，默认排除。

这里把它们合并进 outlook-auto-register 自己的 SQLite 账号库，走同一套运维台 /
Mailbox API 统一管理、收信——**只读**查询 qoderji 的库，从不写回。

## 与本地新注册号的区别：孵化期

本地新注册的号默认进入 `OUTLOOK_INCUBATION_HOURS`（默认 48h）孵化期，期间批量测活 /
保活跳过，避免对新号高频打微软接口触发风控（见 `lifecycle.py`）。外部导入的号早就
存活过一段时间，不需要（也不应该）再走这段冷启动：默认 `skip_incubation=True`，
把 `created_at` 回填到 `OUTLOOK_EXTERNAL_IMPORT_BACKDATE_DAYS`（默认 30）天前，
`lifecycle.enrich_lifecycle_fields` 据此判定 `incubating=False`，运维台 / 保活 /
Mailbox API 立即可用。
"""
from __future__ import annotations

import glob
import logging
import os
import sqlite3
import threading
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Optional

from . import account_store
from . import database as db

logger = logging.getLogger(__name__)

_lock = threading.Lock()

#: qoderji 里代表「原料还能用」的状态；`dead` = OAuth 永久失效，`leased` = 正被某台
#: 注册机租着（我们只读，不影响它的租约，照样能拿来收信）。
DEFAULT_QODERJI_STATUSES: tuple[str, ...] = ("untried", "leased", "consumed")

#: env `QODERJI_EMAIL_DB` 未配置时的默认探测路径（glob）。
DEFAULT_QODERJI_DB_GLOBS: tuple[str, ...] = (
    "/opt/qoderji/data/*.db",
    "/opt/qoderji/*.db",
)

#: 外源号跳过孵化期时，created_at 回填到多少天前。
DEFAULT_BACKDATE_DAYS = 30.0


def _backdate_days() -> float:
    raw = (os.environ.get("OUTLOOK_EXTERNAL_IMPORT_BACKDATE_DAYS") or "").strip()
    if not raw:
        return DEFAULT_BACKDATE_DAYS
    try:
        return max(0.0, float(raw))
    except ValueError:
        return DEFAULT_BACKDATE_DAYS


@dataclass
class ExternalCombo:
    """解析出的一条外部账号 combo。"""

    email: str
    password: str = ""
    client_id: str = ""
    refresh_token: str = ""
    login_client_id: str = ""
    login_refresh_token: str = ""
    recovery_email: str = ""
    recovery_password: str = ""
    raw_line: str = ""
    extra: dict[str, Any] = field(default_factory=dict)

    def combo4(self) -> str:
        return "----".join([self.email, self.password, self.client_id, self.refresh_token])

    def combo6_dual(self) -> str:
        if not self.login_refresh_token:
            return ""
        return "----".join([
            self.email, self.password, self.client_id, self.refresh_token,
            self.login_client_id, self.login_refresh_token,
        ])

    def combo6_recovery(self) -> str:
        if not self.recovery_email:
            return ""
        return "----".join([
            self.email, self.password, self.client_id, self.refresh_token,
            self.recovery_email, self.recovery_password,
        ])


def parse_combo_line(line: str) -> Optional[ExternalCombo]:
    """解析一行 combo。

    - 4 段：``email----password----client_id----refresh_token``（qoderji 的标准形态，
      也是 Graph 四段 combo）。
    - 6 段：末两段固定视为登录令牌 ``login_client_id----login_refresh_token``
      （与运维台既有的 `/api/accounts/import` 粘贴导入语义一致）。
    - 少于 4 段但有邮箱（裸邮箱 / 邮箱+密码）：仍返回，供上层按「缺 token」判定跳过。
    - 分隔符默认 ``----``；没有 ``----`` 时兜底按 ``|`` / tab / 逗号拆一次首段取邮箱。
    """
    s = (line or "").strip()
    if not s or s.startswith("#"):
        return None
    parts = s.split("----")
    if len(parts) < 2:
        for sep in ("|", "\t", ","):
            if sep in s:
                parts = [p.strip() for p in s.split(sep)]
                break
    email = parts[0].strip().strip("\"'")
    if not email or "@" not in email or email.startswith("@") or email.endswith("@"):
        return None
    combo = ExternalCombo(email=email, raw_line=s)
    if len(parts) >= 2:
        combo.password = parts[1].strip()
    if len(parts) >= 4:
        combo.client_id = parts[2].strip()
        combo.refresh_token = parts[3].strip()
    if len(parts) >= 6 and parts[4].strip() and parts[5].strip():
        combo.login_client_id = parts[4].strip()
        combo.login_refresh_token = parts[5].strip()
    return combo


def parse_combo_text(text: str) -> tuple[list[ExternalCombo], dict[str, int]]:
    """按行解析一段 combo 文本，返回 (解析出的 combo 列表, {invalid: N})。"""
    combos: list[ExternalCombo] = []
    invalid = 0
    for raw in (text or "").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        combo = parse_combo_line(line)
        if combo is None:
            invalid += 1
            continue
        combos.append(combo)
    return combos, {"invalid": invalid}


def import_combos(
    combos: Iterable[ExternalCombo],
    *,
    source: str = "",
    batch_label: str = "",
    skip_incubation: bool = False,
    backdate_days: Optional[float] = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """批量 upsert 外部 combo 进账号库。按邮箱去重（不覆盖已存在账号的任何字段）。

    返回 ``{imported, duplicate, invalid, six_seg, emails}``；``dry_run=True`` 时只统计
    不写库。
    """
    db.ensure_initialized()
    now = datetime.now()
    if skip_incubation:
        days = _backdate_days() if backdate_days is None else max(0.0, float(backdate_days))
        created_at = (now - timedelta(days=days)).isoformat(timespec="seconds")
    else:
        created_at = now.isoformat(timespec="seconds")
    updated_at = now.isoformat(timespec="seconds")

    batch_id = f"external:{source}" if source else "external"
    label = batch_label or (f"外源导入:{source}" if source else "外源导入")
    legacy_source = f"external:{source}" if source else "external_import"
    tags = [f"src:{source}"] if source else []

    imported = duplicate = invalid = six_seg = 0
    imported_emails: list[str] = []

    with _lock:
        conn = db.connect()
        try:
            existing = {
                (r["email"] or "").lower()
                for r in conn.execute("SELECT email FROM accounts")
            }
            seen: set[str] = set()
            for combo in combos:
                email = (combo.email or "").strip()
                if not email or "@" not in email:
                    invalid += 1
                    continue
                key = email.lower()
                if key in existing or key in seen:
                    duplicate += 1
                    continue
                if not combo.refresh_token:
                    invalid += 1
                    continue
                seen.add(key)
                row = {
                    "email": email,
                    "password": combo.password,
                    "client_id": combo.client_id,
                    "refresh_token": combo.refresh_token,
                    "login_client_id": combo.login_client_id,
                    "login_refresh_token": combo.login_refresh_token,
                    "recovery_email": combo.recovery_email,
                    "recovery_password": combo.recovery_password,
                    "combo": combo.combo4(),
                    "combo_dual": combo.combo6_dual(),
                    "combo_recovery": combo.combo6_recovery(),
                    "success": True,
                    "created_at": created_at,
                    "updated_at": updated_at,
                    "batch_id": batch_id,
                    "batch_label": label,
                    "legacy_source": legacy_source,
                    "extra": dict(combo.extra),
                }
                if not dry_run:
                    account_store.upsert_account_dict(conn, row)
                    if tags:
                        account_store.upsert_meta_dict(conn, email, {
                            "tags": list(tags),
                            "updated_at": updated_at,
                        })
                imported += 1
                imported_emails.append(email)
                if combo.login_refresh_token:
                    six_seg += 1
            if not dry_run and imported:
                conn.commit()
        finally:
            conn.close()

    return {
        "imported": imported,
        "duplicate": duplicate,
        "invalid": invalid,
        "six_seg": six_seg,
        "emails": imported_emails,
        "source": source,
        "batch_label": label,
        "skip_incubation": skip_incubation,
        "dry_run": dry_run,
    }


def import_text(
    text: str,
    *,
    source: str = "",
    batch_label: str = "",
    skip_incubation: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """解析一段粘贴的 combo 文本并导入，供 webapp `/api/accounts/import` 直接调用。"""
    combos, parse_stats = parse_combo_text(text)
    result = import_combos(
        combos,
        source=source,
        batch_label=batch_label,
        skip_incubation=skip_incubation,
        dry_run=dry_run,
    )
    result["invalid"] += parse_stats.get("invalid", 0)
    return result


# ---------------------------------------------------------------------------
# qoderji email_inventory：只读拉取
# ---------------------------------------------------------------------------


def qoderji_db_candidates(explicit: Optional[str] = None) -> list[Path]:
    """列出可用的 qoderji sqlite 路径。

    ``explicit`` 优先（一般来自 ``QODERJI_EMAIL_DB`` 或调用参数），支持逗号分隔多个
    路径 / glob；留空则按 `DEFAULT_QODERJI_DB_GLOBS` 探测。只返回真实存在的文件。
    """
    raw = explicit if explicit is not None else os.environ.get("QODERJI_EMAIL_DB", "")
    raw = (raw or "").strip()
    patterns = [p.strip() for p in raw.split(",") if p.strip()] if raw else list(DEFAULT_QODERJI_DB_GLOBS)

    found: list[Path] = []
    seen: set[str] = set()
    for pattern in patterns:
        hits = sorted(glob.glob(pattern)) or ([pattern] if os.path.exists(pattern) else [])
        for hit in hits:
            p = Path(hit)
            if not p.exists() or not p.is_file():
                continue
            key = str(p.resolve())
            if key in seen:
                continue
            seen.add(key)
            found.append(p)
    return found


def _qoderji_table_exists(conn: sqlite3.Connection) -> bool:
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='email_inventory'"
    ).fetchone()
    return row is not None


def fetch_qoderji_combos(
    db_path: os.PathLike[str] | str,
    *,
    statuses: Iterable[str] = DEFAULT_QODERJI_STATUSES,
    batch_id: Optional[str] = None,
    limit: Optional[int] = None,
) -> tuple[list[ExternalCombo], dict[str, Any]]:
    """只读连接一个 qoderji sqlite 库，拉出 `email_inventory` 里的 combo。

    从不写这个库（``mode=ro`` URI 连接），可以在 qoderji 正常运行、持续写入时安全跑。
    """
    path = Path(db_path)
    stats: dict[str, Any] = {
        "db_path": str(path),
        "scanned": 0,
        "parsed": 0,
        "invalid": 0,
        "by_status": {},
    }
    combos: list[ExternalCombo] = []
    if not path.exists():
        stats["error"] = "文件不存在"
        return combos, stats

    statuses = tuple(statuses) if statuses else ()
    uri = f"file:{path.resolve()}?mode=ro"
    try:
        conn = sqlite3.connect(uri, uri=True, timeout=10)
    except sqlite3.OperationalError as exc:
        stats["error"] = f"打开失败: {exc}"
        return combos, stats
    conn.row_factory = sqlite3.Row
    try:
        if not _qoderji_table_exists(conn):
            stats["error"] = "email_inventory 表不存在"
            return combos, stats

        clauses: list[str] = []
        params: list[Any] = []
        if statuses:
            placeholders = ",".join("?" for _ in statuses)
            clauses.append(f"status IN ({placeholders})")
            params.extend(statuses)
        if batch_id:
            clauses.append("batch_id = ?")
            params.append(batch_id)
        query = "SELECT email, raw, status, batch_id, source, added_at FROM email_inventory"
        if clauses:
            query += " WHERE " + " AND ".join(clauses)
        query += " ORDER BY added_at ASC"
        if limit:
            query += " LIMIT ?"
            params.append(int(limit))

        for row in conn.execute(query, params):
            stats["scanned"] += 1
            st = row["status"] or ""
            stats["by_status"][st] = stats["by_status"].get(st, 0) + 1
            raw_line = row["raw"] or row["email"] or ""
            combo = parse_combo_line(raw_line)
            if combo is None or not combo.refresh_token:
                stats["invalid"] += 1
                continue
            combo.extra.update({
                "qoderji_status": st,
                "qoderji_batch_id": row["batch_id"] or "",
                "qoderji_source": row["source"] or "",
                "qoderji_added_at": row["added_at"],
            })
            combos.append(combo)
            stats["parsed"] += 1
    finally:
        conn.close()
    return combos, stats


def import_from_qoderji(
    *,
    db_path: Optional[str] = None,
    statuses: Iterable[str] = DEFAULT_QODERJI_STATUSES,
    batch_id: Optional[str] = None,
    limit: Optional[int] = None,
    batch_label: str = "",
    skip_incubation: bool = True,
    dry_run: bool = False,
) -> dict[str, Any]:
    """从 qoderji `email_inventory` 拉取 + 导入的一站式入口。

    ``db_path`` 留空时按 `QODERJI_EMAIL_DB` / 默认路径探测；同一台机可能存在多个候选
    （如 `qoderji.db` 的旧拷贝 + `data/cards.db` 的实际库），会依次扫描并合并去重。
    """
    statuses = tuple(statuses) if statuses else DEFAULT_QODERJI_STATUSES
    candidates = qoderji_db_candidates(db_path)
    if not candidates:
        return {
            "ok": False,
            "error": "找不到 qoderji 邮箱库：QODERJI_EMAIL_DB 未配置，默认路径下也没有 *.db",
            "db_files": [],
            "imported": 0,
            "duplicate": 0,
            "invalid": 0,
            "six_seg": 0,
        }

    all_combos: list[ExternalCombo] = []
    per_db_stats: list[dict[str, Any]] = []
    remaining = limit
    for path in candidates:
        combos, stats = fetch_qoderji_combos(
            path, statuses=statuses, batch_id=batch_id, limit=remaining,
        )
        per_db_stats.append(stats)
        all_combos.extend(combos)
        if remaining is not None:
            remaining = max(0, remaining - len(combos))
            if remaining <= 0:
                break

    result = import_combos(
        all_combos,
        source="qoderji",
        batch_label=batch_label or "qoderji导入",
        skip_incubation=skip_incubation,
        dry_run=dry_run,
    )
    result["ok"] = True
    result["db_files"] = [str(p) for p in candidates]
    result["fetch_stats"] = per_db_stats
    result["fetched"] = len(all_combos)
    return result
