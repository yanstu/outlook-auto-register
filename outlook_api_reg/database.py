"""SQLite 统一存储：账号、元数据、批次、代理池、救援日志 + 遗留 JSON 迁移与备份。"""
from __future__ import annotations

import json
import logging
import os
import re
import shutil
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger(__name__)

_lock = threading.Lock()
_initialized = False

SCHEMA_VERSION = 4

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS app_meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS accounts (
    email TEXT PRIMARY KEY COLLATE NOCASE,
    password TEXT NOT NULL DEFAULT '',
    client_id TEXT NOT NULL DEFAULT '',
    refresh_token TEXT NOT NULL DEFAULT '',
    login_client_id TEXT NOT NULL DEFAULT '',
    login_refresh_token TEXT NOT NULL DEFAULT '',
    recovery_email TEXT NOT NULL DEFAULT '',
    recovery_password TEXT NOT NULL DEFAULT '',
    combo TEXT NOT NULL DEFAULT '',
    combo_dual TEXT NOT NULL DEFAULT '',
    combo_recovery TEXT NOT NULL DEFAULT '',
    redirect_url TEXT NOT NULL DEFAULT '',
    auth_status TEXT NOT NULL DEFAULT '',
    proofs_method TEXT NOT NULL DEFAULT '',
    proofs_satisfied TEXT NOT NULL DEFAULT '',
    login_status TEXT NOT NULL DEFAULT '',
    login_fail_reason TEXT NOT NULL DEFAULT '',
    dual_requested INTEGER NOT NULL DEFAULT 0,
    dual_ok INTEGER NOT NULL DEFAULT 0,
    success INTEGER NOT NULL DEFAULT 1,
    error TEXT NOT NULL DEFAULT '',
    extra_json TEXT NOT NULL DEFAULT '{}',
    batch_id TEXT NOT NULL DEFAULT '',
    batch_no INTEGER,
    batch_label TEXT NOT NULL DEFAULT '',
    rescue_count INTEGER NOT NULL DEFAULT 0,
    last_rescue_at TEXT NOT NULL DEFAULT '',
    last_rescue_ok INTEGER,
    last_rescue_reason TEXT NOT NULL DEFAULT '',
    rescued_at TEXT NOT NULL DEFAULT '',
    rescued_scope TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL DEFAULT '',
    last_alive_at TEXT NOT NULL DEFAULT '',
    legacy_source TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS account_meta (
    email TEXT PRIMARY KEY COLLATE NOCASE,
    note TEXT NOT NULL DEFAULT '',
    tags_json TEXT NOT NULL DEFAULT '[]',
    verify_json TEXT NOT NULL DEFAULT '',
    combo_dual_meta TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (email) REFERENCES accounts(email) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS register_jobs (
    id TEXT PRIMARY KEY,
    batch_no INTEGER,
    batch_label TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '',
    count INTEGER NOT NULL DEFAULT 0,
    concurrency INTEGER NOT NULL DEFAULT 1,
    token_mode TEXT NOT NULL DEFAULT '',
    dry_run INTEGER NOT NULL DEFAULT 0,
    ok_count INTEGER NOT NULL DEFAULT 0,
    fail_count INTEGER NOT NULL DEFAULT 0,
    emails_json TEXT NOT NULL DEFAULT '[]',
    params_json TEXT NOT NULL DEFAULT '{}'
);

CREATE TABLE IF NOT EXISTS proxy_settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS proxies (
    id TEXT PRIMARY KEY,
    label TEXT NOT NULL,
    provider TEXT NOT NULL DEFAULT '',
    country TEXT NOT NULL DEFAULT '',
    template TEXT NOT NULL,
    enabled INTEGER NOT NULL DEFAULT 1,
    status TEXT NOT NULL DEFAULT 'unknown',
    last_check_at TEXT NOT NULL DEFAULT '',
    last_check_msg TEXT NOT NULL DEFAULT '',
    exit_ip TEXT NOT NULL DEFAULT '',
    assigned_count INTEGER NOT NULL DEFAULT 0,
    success_count INTEGER NOT NULL DEFAULT 0,
    fail_count INTEGER NOT NULL DEFAULT 0,
    check_count INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT '',
    sort_order INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS proxy_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    proxy_id TEXT NOT NULL DEFAULT '',
    provider TEXT NOT NULL DEFAULT '',
    country TEXT NOT NULL DEFAULT '',
    reg_country TEXT NOT NULL DEFAULT '',
    purpose TEXT NOT NULL DEFAULT 'register',
    success INTEGER NOT NULL DEFAULT 0,
    email TEXT NOT NULL DEFAULT '',
    error TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS proxy_bindings (
    email TEXT PRIMARY KEY COLLATE NOCASE,
    proxy_id TEXT NOT NULL,
    resolved TEXT NOT NULL,
    assigned_at TEXT NOT NULL,
    purpose TEXT NOT NULL DEFAULT 'register'
);

CREATE TABLE IF NOT EXISTS rescue_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL COLLATE NOCASE,
    ok INTEGER NOT NULL DEFAULT 0,
    reason TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS api_principals (
    id TEXT PRIMARY KEY,
    kind TEXT NOT NULL DEFAULT 'service',
    name TEXT NOT NULL DEFAULT '',
    secret_hash TEXT NOT NULL DEFAULT '',
    scopes TEXT NOT NULL DEFAULT '',
    enabled INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL DEFAULT '',
    expires_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS api_grants (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    principal_id TEXT NOT NULL,
    email TEXT NOT NULL DEFAULT '' COLLATE NOCASE,
    scopes_override TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS api_sessions (
    id TEXT PRIMARY KEY,
    token_hash TEXT NOT NULL DEFAULT '',
    principal_id TEXT NOT NULL DEFAULT '',
    email TEXT NOT NULL DEFAULT '' COLLATE NOCASE,
    expires_at TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS api_audit (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts TEXT NOT NULL DEFAULT '',
    principal_id TEXT NOT NULL DEFAULT '',
    method TEXT NOT NULL DEFAULT '',
    path TEXT NOT NULL DEFAULT '',
    email TEXT NOT NULL DEFAULT '',
    status INTEGER NOT NULL DEFAULT 0,
    detail TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_accounts_batch ON accounts(batch_label);
CREATE INDEX IF NOT EXISTS idx_accounts_created ON accounts(created_at);
CREATE INDEX IF NOT EXISTS idx_proxies_enabled ON proxies(enabled);
CREATE INDEX IF NOT EXISTS idx_proxies_status ON proxies(status);
CREATE INDEX IF NOT EXISTS idx_proxies_provider ON proxies(provider);
CREATE INDEX IF NOT EXISTS idx_proxies_country ON proxies(country);
CREATE INDEX IF NOT EXISTS idx_bindings_proxy_id ON proxy_bindings(proxy_id);
CREATE INDEX IF NOT EXISTS idx_proxy_events_provider ON proxy_events(provider, country);
CREATE INDEX IF NOT EXISTS idx_proxy_events_proxy_id ON proxy_events(proxy_id);
CREATE INDEX IF NOT EXISTS idx_rescue_events_email ON rescue_events(email);
CREATE UNIQUE INDEX IF NOT EXISTS idx_api_grants_unique ON api_grants(principal_id, email);
CREATE INDEX IF NOT EXISTS idx_api_sessions_principal ON api_sessions(principal_id);
CREATE INDEX IF NOT EXISTS idx_api_sessions_email ON api_sessions(email);
CREATE INDEX IF NOT EXISTS idx_api_audit_ts ON api_audit(ts);
CREATE INDEX IF NOT EXISTS idx_api_audit_principal ON api_audit(principal_id);
"""

# v4 新增的 Mailbox API 表：老库升级时单独建一次，之后交给 SCHEMA_SQL 幂等维护
MAILBOX_API_SQL = """
CREATE TABLE IF NOT EXISTS api_principals (
    id TEXT PRIMARY KEY,
    kind TEXT NOT NULL DEFAULT 'service',
    name TEXT NOT NULL DEFAULT '',
    secret_hash TEXT NOT NULL DEFAULT '',
    scopes TEXT NOT NULL DEFAULT '',
    enabled INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL DEFAULT '',
    expires_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS api_grants (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    principal_id TEXT NOT NULL,
    email TEXT NOT NULL DEFAULT '' COLLATE NOCASE,
    scopes_override TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS api_sessions (
    id TEXT PRIMARY KEY,
    token_hash TEXT NOT NULL DEFAULT '',
    principal_id TEXT NOT NULL DEFAULT '',
    email TEXT NOT NULL DEFAULT '' COLLATE NOCASE,
    expires_at TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS api_audit (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts TEXT NOT NULL DEFAULT '',
    principal_id TEXT NOT NULL DEFAULT '',
    method TEXT NOT NULL DEFAULT '',
    path TEXT NOT NULL DEFAULT '',
    email TEXT NOT NULL DEFAULT '',
    status INTEGER NOT NULL DEFAULT 0,
    detail TEXT NOT NULL DEFAULT ''
);
"""

_MAILBOX_API_COLUMNS = {
    "api_principals": {
        "kind": "TEXT NOT NULL DEFAULT 'service'",
        "name": "TEXT NOT NULL DEFAULT ''",
        "secret_hash": "TEXT NOT NULL DEFAULT ''",
        "scopes": "TEXT NOT NULL DEFAULT ''",
        "enabled": "INTEGER NOT NULL DEFAULT 1",
        "created_at": "TEXT NOT NULL DEFAULT ''",
        "expires_at": "TEXT NOT NULL DEFAULT ''",
    },
    "api_grants": {
        "principal_id": "TEXT NOT NULL DEFAULT ''",
        "email": "TEXT NOT NULL DEFAULT ''",
        "scopes_override": "TEXT NOT NULL DEFAULT ''",
        "created_at": "TEXT NOT NULL DEFAULT ''",
    },
    "api_sessions": {
        "token_hash": "TEXT NOT NULL DEFAULT ''",
        "principal_id": "TEXT NOT NULL DEFAULT ''",
        "email": "TEXT NOT NULL DEFAULT ''",
        "expires_at": "TEXT NOT NULL DEFAULT ''",
        "created_at": "TEXT NOT NULL DEFAULT ''",
    },
    "api_audit": {
        "ts": "TEXT NOT NULL DEFAULT ''",
        "principal_id": "TEXT NOT NULL DEFAULT ''",
        "method": "TEXT NOT NULL DEFAULT ''",
        "path": "TEXT NOT NULL DEFAULT ''",
        "email": "TEXT NOT NULL DEFAULT ''",
        "status": "INTEGER NOT NULL DEFAULT 0",
        "detail": "TEXT NOT NULL DEFAULT ''",
    },
}

_FN_TS = re.compile(r"_(\d{8})_(\d{6})\.json$")


def db_path() -> Path:
    env = os.environ.get("OUTLOOK_DB_PATH", "").strip()
    if env:
        return Path(env).expanduser()
    root = Path(__file__).resolve().parent.parent / "accounts"
    return root / "outlook.db"


def backup_dir() -> Path:
    return db_path().parent / "backups"


def storage_backend() -> str:
    return "sqlite"


def connect() -> sqlite3.Connection:
    path = db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def _app_get(conn: sqlite3.Connection, key: str, default: str = "") -> str:
    row = conn.execute("SELECT value FROM app_meta WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default


def _app_set(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute(
        "INSERT INTO app_meta(key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value),
    )


_SETTING_PREFIX = "setting:"


def get_setting(key: str, default: str = "") -> str:
    """读取一条应用设置（DB 为主存储）。key 不含前缀，内部落在 app_meta 的 setting:<key>。"""
    ensure_initialized()
    conn = connect()
    try:
        return _app_get(conn, _SETTING_PREFIX + key, default)
    finally:
        conn.close()


def set_setting(key: str, value: str) -> None:
    """写入一条应用设置（DB 为主存储）。value 为空串等于清空。"""
    ensure_initialized()
    conn = connect()
    try:
        _app_set(conn, _SETTING_PREFIX + key, value or "")
        conn.commit()
    finally:
        conn.close()


def apply_schema(conn: sqlite3.Connection) -> None:
    _migrate_schema_v3(conn)
    _migrate_schema_v4(conn)
    conn.executescript(SCHEMA_SQL)
    _app_set(conn, "schema_version", str(SCHEMA_VERSION))
    conn.commit()


def _migrate_schema_v3(conn: sqlite3.Connection) -> None:
    """增量迁移：旧库补 country 列、proxy_events 表。"""
    cols = {r[1] for r in conn.execute("PRAGMA table_info(proxies)")}
    if cols and "country" not in cols:
        conn.execute("ALTER TABLE proxies ADD COLUMN country TEXT NOT NULL DEFAULT ''")
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS proxy_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proxy_id TEXT NOT NULL DEFAULT '',
            provider TEXT NOT NULL DEFAULT '',
            country TEXT NOT NULL DEFAULT '',
            reg_country TEXT NOT NULL DEFAULT '',
            purpose TEXT NOT NULL DEFAULT 'register',
            success INTEGER NOT NULL DEFAULT 0,
            email TEXT NOT NULL DEFAULT '',
            error TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL DEFAULT ''
        );
        CREATE INDEX IF NOT EXISTS idx_proxy_events_provider ON proxy_events(provider, country);
        CREATE INDEX IF NOT EXISTS idx_proxy_events_proxy_id ON proxy_events(proxy_id);
        """
    )
    if not cols:
        return
    from .proxy_utils import infer_country_from_template

    for row in conn.execute("SELECT id, template, country FROM proxies WHERE country='' OR country IS NULL"):
        cc = infer_country_from_template(row["template"] or "")
        if cc:
            conn.execute("UPDATE proxies SET country=? WHERE id=?", (cc, row["id"]))


def _migrate_schema_v4(conn: sqlite3.Connection) -> None:
    """增量迁移：补 Mailbox API 的 principal / grant / session / audit 四张表与缺失列。"""
    conn.executescript(MAILBOX_API_SQL)
    for table, columns in _MAILBOX_API_COLUMNS.items():
        existing = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
        if not existing:
            continue
        for name, ddl in columns.items():
            if name not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


def backup_database(*, tag: str = "") -> Path:
    """热备份 SQLite 到 accounts/backups/。"""
    dest = backup_database_safe(tag=tag)
    if dest is None:
        raise FileNotFoundError("备份失败或数据库不存在")
    return dest


def backup_database_safe(*, tag: str = "") -> Optional[Path]:
    try:
        with _lock:
            src = db_path()
            if not src.exists():
                return None
            dest_dir = backup_dir()
            dest_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            suffix = f"_{tag}" if tag else ""
            dest = dest_dir / f"outlook_{stamp}{suffix}.db"
            src_conn = sqlite3.connect(str(src))
            dest_conn = sqlite3.connect(str(dest))
            try:
                src_conn.backup(dest_conn)
                dest_conn.commit()
            finally:
                dest_conn.close()
                src_conn.close()
            meta_conn = connect()
            try:
                _app_set(meta_conn, "last_backup_at", datetime.now().isoformat(timespec="seconds"))
                meta_conn.commit()
            finally:
                meta_conn.close()
            # prune old backups (keep 30)
            backups = sorted(dest_dir.glob("outlook_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)
            for old in backups[30:]:
                try:
                    old.unlink()
                except Exception:  # noqa: BLE001
                    pass
            logger.info("数据库已备份: %s", dest)
            return dest
    except Exception as exc:  # noqa: BLE001
        logger.warning("数据库备份失败: %s", exc)
        return None


def _infer_created_at(fp: Path, data: dict[str, Any]) -> str:
    if data.get("created_at"):
        return str(data["created_at"])
    m = _FN_TS.search(fp.name)
    if m:
        d, t = m.groups()
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}T{t[:2]}:{t[2:4]}:{t[4:6]}"
    return ""


def _archive_legacy(accounts_dir: Path, names: list[str]) -> None:
    if not names:
        return
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arc = accounts_dir / ".archive" / stamp
    arc.mkdir(parents=True, exist_ok=True)
    for name in names:
        src = accounts_dir / name
        if src.exists() and src.is_file():
            try:
                shutil.move(str(src), str(arc / name))
            except Exception as exc:  # noqa: BLE001
                logger.warning("归档 %s 失败: %s", name, exc)


def migrate_legacy_files(accounts_dir: Path) -> dict[str, int]:
    """一次性把 JSON/txt 遗留数据迁入 SQLite，并归档源文件。"""
    stats = {"accounts": 0, "meta": 0, "jobs": 0, "rescue_events": 0, "txt_lines": 0}
    accounts_dir = Path(accounts_dir)
    accounts_dir.mkdir(parents=True, exist_ok=True)

    from .account_persist import merge_account_row

    conn = connect()
    try:
        apply_schema(conn)
        if _app_get(conn, "legacy_files_migrated") == "1":
            return stats
        backup_database_safe(tag="pre-migrate")

        items: dict[str, dict[str, Any]] = {}
        archived: list[str] = []

        meta_file = accounts_dir / "webapp_meta.json"
        meta: dict[str, Any] = {}
        if meta_file.exists():
            try:
                meta = json.loads(meta_file.read_text(encoding="utf-8"))
            except Exception:  # noqa: BLE001
                meta = {}

        for fp in sorted(accounts_dir.glob("*.json")):
            if fp.name in ("webapp_meta.json", "webapp_jobs.json", "proxy_pool.json"):
                continue
            try:
                data = json.loads(fp.read_text(encoding="utf-8"))
            except Exception:  # noqa: BLE001
                continue
            if not isinstance(data, dict):
                continue
            email = (data.get("email") or "").strip()
            if not email:
                continue
            created_at = _infer_created_at(fp, data)
            row = _account_dict_from_json(data, email, created_at, fp.name)
            if email in items:
                merge_account_row(items[email], row, source=fp.name)
            else:
                items[email] = row
            archived.append(fp.name)

        for fname in ("accounts_recovery.txt", "accounts.txt", "accounts_dual.txt"):
            fp = accounts_dir / fname
            if not fp.exists():
                continue
            for line in fp.read_text(encoding="utf-8", errors="replace").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "----" not in line:
                    continue
                parts = line.split("----")
                email = parts[0].strip()
                if not email:
                    continue
                if fname == "accounts_recovery.txt" and len(parts) >= 6:
                    row = {
                        "email": email,
                        "password": parts[1],
                        "client_id": parts[2],
                        "refresh_token": parts[3],
                        "recovery_email": parts[4],
                        "recovery_password": parts[5],
                        "combo": "----".join(parts[:4]),
                        "combo_recovery": line,
                        "has_recovery": True,
                        "legacy_source": fname,
                    }
                elif fname == "accounts_dual.txt" and len(parts) >= 6:
                    row = {
                        "email": email,
                        "password": parts[1],
                        "client_id": parts[2],
                        "refresh_token": parts[3],
                        "combo": "----".join(parts[:4]),
                        "combo_dual": line,
                        "login_token": True,
                        "legacy_source": fname,
                    }
                elif len(parts) >= 4:
                    if email in items:
                        continue
                    row = {
                        "email": email,
                        "password": parts[1],
                        "client_id": parts[2],
                        "refresh_token": parts[3],
                        "combo": line,
                        "legacy_source": fname,
                    }
                else:
                    continue
                if email in items:
                    merge_account_row(items[email], row, source=fname)
                else:
                    items[email] = row
                stats["txt_lines"] += 1

        from .account_store import upsert_account_dict, upsert_meta_dict

        for email, row in items.items():
            upsert_account_dict(conn, row)
            stats["accounts"] += 1
            m = meta.get(email) or {}
            if m:
                upsert_meta_dict(conn, email, m)
                stats["meta"] += 1

        for email, m in meta.items():
            if email not in items and m:
                upsert_account_dict(conn, {
                    "email": email,
                    "created_at": datetime.now().isoformat(timespec="seconds"),
                    "updated_at": datetime.now().isoformat(timespec="seconds"),
                    "legacy_source": "meta-only",
                })
                upsert_meta_dict(conn, email, m)
                stats["meta"] += 1

        jobs_file = accounts_dir / "webapp_jobs.json"
        if jobs_file.exists():
            try:
                jobs = json.loads(jobs_file.read_text(encoding="utf-8"))
                if isinstance(jobs, list):
                    for rec in jobs:
                        conn.execute(
                            """INSERT OR REPLACE INTO register_jobs(
                                id, batch_no, batch_label, created_at, status, count, concurrency,
                                token_mode, dry_run, ok_count, fail_count, emails_json, params_json
                            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (
                                rec.get("id") or "",
                                rec.get("batch_no"),
                                rec.get("batch_label") or "",
                                rec.get("created_at") or "",
                                rec.get("status") or "",
                                int(rec.get("count") or 0),
                                int(rec.get("concurrency") or 1),
                                rec.get("token_mode") or "",
                                1 if rec.get("dry_run") else 0,
                                int(rec.get("ok_count") or 0),
                                int(rec.get("fail_count") or 0),
                                json.dumps(rec.get("emails") or [], ensure_ascii=False),
                                json.dumps({}, ensure_ascii=False),
                            ),
                        )
                        stats["jobs"] += 1
                    archived.append(jobs_file.name)
            except Exception as exc:  # noqa: BLE001
                logger.warning("迁移 jobs 失败: %s", exc)

        rescue_log = accounts_dir / "rescue_results.jsonl"
        if rescue_log.exists():
            for line in rescue_log.read_text(encoding="utf-8", errors="replace").splitlines():
                line = line.strip()
                if not line:
                    continue
                try:
                    ev = json.loads(line)
                except Exception:  # noqa: BLE001
                    continue
                em = (ev.get("email") or "").strip()
                if not em:
                    continue
                conn.execute(
                    "INSERT INTO rescue_events(email, ok, reason, created_at) VALUES (?,?,?,?)",
                    (
                        em.lower(),
                        1 if ev.get("ok") else 0,
                        str(ev.get("reason") or ev.get("message") or "")[:240],
                        str(ev.get("ts") or ev.get("at") or datetime.now().isoformat()),
                    ),
                )
                stats["rescue_events"] += 1

        _migrate_proxy_json(conn, accounts_dir)

        _app_set(conn, "legacy_files_migrated", "1")
        _app_set(conn, "legacy_migrated_at", datetime.now().isoformat(timespec="seconds"))
        conn.commit()
        _archive_legacy(accounts_dir, archived)
        if meta_file.exists() and meta_file.name not in archived:
            _archive_legacy(accounts_dir, [meta_file.name])
    finally:
        conn.close()
    logger.info("遗留文件迁移完成: %s", stats)
    return stats


def _migrate_proxy_json(conn: sqlite3.Connection, accounts_dir: Path) -> None:
    legacy = accounts_dir / "proxy_pool.json"
    if not legacy.exists():
        return
    n = conn.execute("SELECT COUNT(*) AS c FROM proxies").fetchone()["c"]
    if n:
        return
    try:
        data = json.loads(legacy.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return
    for key, val in (data.get("settings") or {}).items():
        conn.execute(
            "INSERT OR IGNORE INTO proxy_settings(key, value) VALUES (?, ?)",
            (key, json.dumps(val)),
        )
    for i, p in enumerate(data.get("proxies") or []):
        from .proxy_utils import infer_country_from_template

        stats = p.get("stats") or {}
        conn.execute(
            """INSERT INTO proxies(
                id, label, provider, country, template, enabled, status, last_check_at, last_check_msg,
                exit_ip, assigned_count, success_count, fail_count, check_count, created_at, sort_order
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                p.get("id") or "",
                p.get("label") or "",
                p.get("provider") or "",
                p.get("country") or infer_country_from_template(p.get("template") or ""),
                p.get("template") or "",
                1 if p.get("enabled", True) else 0,
                p.get("status") or "unknown",
                p.get("last_check_at") or "",
                p.get("last_check_msg") or "",
                p.get("exit_ip") or "",
                int(stats.get("assigned") or 0),
                int(stats.get("success") or 0),
                int(stats.get("fail") or 0),
                int(stats.get("checks") or 0),
                p.get("created_at") or datetime.now().isoformat(timespec="seconds"),
                i,
            ),
        )
    for email, b in (data.get("bindings") or {}).items():
        conn.execute(
            """INSERT OR REPLACE INTO proxy_bindings(email, proxy_id, resolved, assigned_at, purpose)
               VALUES (?,?,?,?,?)""",
            (
                email.strip().lower(),
                b.get("proxy_id") or "",
                b.get("resolved") or "",
                b.get("assigned_at") or "",
                b.get("purpose") or "register",
            ),
        )
    try:
        legacy.replace(legacy.with_suffix(".json.migrated"))
    except Exception:  # noqa: BLE001
        pass


def _account_dict_from_json(
    data: dict[str, Any],
    email: str,
    created_at: str,
    source: str,
) -> dict[str, Any]:
    return {
        "email": email,
        "password": data.get("password", ""),
        "client_id": data.get("client_id", ""),
        "refresh_token": data.get("refresh_token", ""),
        "login_client_id": data.get("login_client_id", ""),
        "login_refresh_token": data.get("login_refresh_token", ""),
        "recovery_email": data.get("recovery_email", ""),
        "recovery_password": data.get("recovery_password", ""),
        "combo": data.get("combo") or "",
        "combo_dual": data.get("combo_dual", ""),
        "combo_recovery": data.get("combo_recovery", ""),
        "redirect_url": data.get("redirect_url", ""),
        "auth_status": data.get("auth_status", ""),
        "proofs_method": data.get("proofs_method", ""),
        "proofs_satisfied": data.get("proofs_satisfied", ""),
        "login_status": data.get("login_status", ""),
        "login_fail_reason": data.get("login_fail_reason", ""),
        "dual_requested": bool(data.get("dual_requested")),
        "dual_ok": bool(data.get("dual_ok")),
        "success": bool(data.get("success", True)),
        "error": data.get("error", ""),
        "extra": data.get("extra") or {},
        "batch_id": data.get("batch_id", ""),
        "batch_no": data.get("batch_no"),
        "batch_label": data.get("batch_label", ""),
        "rescue_count": int(data.get("rescue_count") or 0),
        "last_rescue_at": data.get("last_rescue_at", ""),
        "last_rescue_ok": data.get("last_rescue_ok"),
        "last_rescue_reason": data.get("last_rescue_reason", ""),
        "rescued_at": data.get("rescued_at", ""),
        "rescued_scope": data.get("rescued_scope", ""),
        "created_at": data.get("created_at") or created_at,
        "updated_at": data.get("updated_at", ""),
        "last_alive_at": data.get("last_alive_at", ""),
        "legacy_source": source,
        "has_token": bool(data.get("refresh_token")),
        "login_token": bool(data.get("combo_dual") or data.get("login_refresh_token")),
        "has_recovery": bool(
            data.get("combo_recovery")
            or (data.get("recovery_email") and data.get("recovery_password"))
        ),
    }


def ensure_initialized(accounts_dir: Optional[Path] = None) -> None:
    global _initialized
    with _lock:
        if _initialized:
            return
    adir = accounts_dir or db_path().parent
    conn = connect()
    try:
        apply_schema(conn)
    finally:
        conn.close()
    migrate_legacy_files(adir)
    with _lock:
        _initialized = True


def db_status() -> dict[str, Any]:
    ensure_initialized()
    path = db_path()
    info: dict[str, Any] = {
        "backend": storage_backend(),
        "path": str(path),
        "exists": path.exists(),
        "size_bytes": path.stat().st_size if path.exists() else 0,
        "backup_dir": str(backup_dir()),
        "last_backup_at": "",
        "legacy_migrated": False,
        "schema_version": SCHEMA_VERSION,
    }
    if not path.exists():
        return info
    conn = connect()
    try:
        info["last_backup_at"] = _app_get(conn, "last_backup_at")
        info["legacy_migrated"] = _app_get(conn, "legacy_files_migrated") == "1"
        info["accounts"] = conn.execute("SELECT COUNT(*) AS c FROM accounts").fetchone()["c"]
        info["proxies"] = conn.execute("SELECT COUNT(*) AS c FROM proxies").fetchone()["c"]
        info["bindings"] = conn.execute("SELECT COUNT(*) AS c FROM proxy_bindings").fetchone()["c"]
        backups = sorted(backup_dir().glob("outlook_*.db"), reverse=True)
        info["backups"] = [p.name for p in backups[:10]]
    finally:
        conn.close()
    return info
