"""账号 / 元数据 / 批次 — SQLite 读写（替代 JSON 文件）。"""
from __future__ import annotations

import json
import logging
import threading
from datetime import datetime
from typing import Any, Optional

from . import database as db
from .account_persist import enrich_register_result, merge_account_row
from .lifecycle import INCUBATING_TAG, enrich_lifecycle_fields
from .models import RegisterResult

logger = logging.getLogger(__name__)
_lock = threading.Lock()


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _row_to_dict(row: Any) -> dict[str, Any]:
    d = dict(row)
    extra = {}
    try:
            extra = json.loads(d.pop("extra_json", "{}") or "{}")
    except Exception:  # noqa: BLE001
        extra = {}
    d["extra"] = extra
    d["has_token"] = bool(d.get("refresh_token"))
    d["login_token"] = bool(d.get("combo_dual") or d.get("login_refresh_token"))
    d["has_recovery"] = bool(
        d.get("combo_recovery") or (d.get("recovery_email") and d.get("recovery_password"))
    )
    d["dual_requested"] = bool(d.get("dual_requested"))
    d["dual_ok"] = bool(d.get("dual_ok"))
    d["success"] = bool(d.get("success"))
    d["source"] = d.get("legacy_source") or "sqlite"
    return d


def upsert_account_dict(conn: Any, row: dict[str, Any]) -> None:
    email = (row.get("email") or "").strip()
    if not email:
        return
    now = _now()
    created = row.get("created_at") or now
    updated = row.get("updated_at") or now
    extra = row.get("extra") or {}
    if not isinstance(extra, dict):
        extra = {}
    conn.execute(
        """INSERT INTO accounts(
            email, password, client_id, refresh_token, login_client_id, login_refresh_token,
            recovery_email, recovery_password, combo, combo_dual, combo_recovery,
            redirect_url, auth_status, proofs_method, proofs_satisfied, login_status,
            login_fail_reason, dual_requested, dual_ok, success, error, extra_json,
            batch_id, batch_no, batch_label, rescue_count, last_rescue_at, last_rescue_ok,
            last_rescue_reason, rescued_at, rescued_scope, created_at, updated_at,
            last_alive_at, legacy_source
        ) VALUES (
            ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
        ) ON CONFLICT(email) DO UPDATE SET
            password=excluded.password,
            client_id=excluded.client_id,
            refresh_token=CASE WHEN excluded.refresh_token!='' THEN excluded.refresh_token ELSE accounts.refresh_token END,
            login_client_id=CASE WHEN excluded.login_client_id!='' THEN excluded.login_client_id ELSE accounts.login_client_id END,
            login_refresh_token=CASE WHEN excluded.login_refresh_token!='' THEN excluded.login_refresh_token ELSE accounts.login_refresh_token END,
            recovery_email=CASE WHEN excluded.recovery_email!='' THEN excluded.recovery_email ELSE accounts.recovery_email END,
            recovery_password=CASE WHEN excluded.recovery_password!='' THEN excluded.recovery_password ELSE accounts.recovery_password END,
            combo=CASE WHEN excluded.combo!='' THEN excluded.combo ELSE accounts.combo END,
            combo_dual=CASE WHEN excluded.combo_dual!='' THEN excluded.combo_dual ELSE accounts.combo_dual END,
            combo_recovery=CASE WHEN excluded.combo_recovery!='' THEN excluded.combo_recovery ELSE accounts.combo_recovery END,
            redirect_url=excluded.redirect_url,
            auth_status=CASE WHEN excluded.auth_status!='' THEN excluded.auth_status ELSE accounts.auth_status END,
            proofs_method=CASE WHEN excluded.proofs_method!='' THEN excluded.proofs_method ELSE accounts.proofs_method END,
            proofs_satisfied=CASE WHEN excluded.proofs_satisfied!='' THEN excluded.proofs_satisfied ELSE accounts.proofs_satisfied END,
            login_status=excluded.login_status,
            login_fail_reason=excluded.login_fail_reason,
            dual_requested=excluded.dual_requested,
            dual_ok=excluded.dual_ok,
            success=excluded.success,
            error=excluded.error,
            extra_json=excluded.extra_json,
            batch_id=CASE WHEN excluded.batch_id!='' THEN excluded.batch_id ELSE accounts.batch_id END,
            batch_no=COALESCE(excluded.batch_no, accounts.batch_no),
            batch_label=CASE WHEN excluded.batch_label!='' THEN excluded.batch_label ELSE accounts.batch_label END,
            rescue_count=MAX(accounts.rescue_count, excluded.rescue_count),
            last_rescue_at=CASE WHEN excluded.last_rescue_at!='' THEN excluded.last_rescue_at ELSE accounts.last_rescue_at END,
            last_rescue_ok=COALESCE(excluded.last_rescue_ok, accounts.last_rescue_ok),
            last_rescue_reason=excluded.last_rescue_reason,
            rescued_at=CASE WHEN excluded.rescued_at!='' THEN excluded.rescued_at ELSE accounts.rescued_at END,
            rescued_scope=CASE WHEN excluded.rescued_scope!='' THEN excluded.rescued_scope ELSE accounts.rescued_scope END,
            updated_at=excluded.updated_at,
            last_alive_at=CASE WHEN excluded.last_alive_at!='' THEN excluded.last_alive_at ELSE accounts.last_alive_at END,
            legacy_source=CASE WHEN excluded.legacy_source!='' THEN excluded.legacy_source ELSE accounts.legacy_source END
        """,
        (
            email,
            row.get("password") or "",
            row.get("client_id") or "",
            row.get("refresh_token") or "",
            row.get("login_client_id") or "",
            row.get("login_refresh_token") or "",
            row.get("recovery_email") or "",
            row.get("recovery_password") or "",
            row.get("combo") or "",
            row.get("combo_dual") or "",
            row.get("combo_recovery") or "",
            row.get("redirect_url") or "",
            row.get("auth_status") or "",
            row.get("proofs_method") or "",
            row.get("proofs_satisfied") or "",
            row.get("login_status") or "",
            row.get("login_fail_reason") or "",
            1 if row.get("dual_requested") else 0,
            1 if row.get("dual_ok") else 0,
            1 if row.get("success", True) else 0,
            row.get("error") or "",
            json.dumps(extra, ensure_ascii=False),
            row.get("batch_id") or "",
            row.get("batch_no"),
            row.get("batch_label") or "",
            int(row.get("rescue_count") or 0),
            row.get("last_rescue_at") or "",
            row.get("last_rescue_ok"),
            row.get("last_rescue_reason") or "",
            row.get("rescued_at") or "",
            row.get("rescued_scope") or "",
            created,
            updated,
            row.get("last_alive_at") or "",
            row.get("legacy_source") or row.get("source") or "",
        ),
    )


def upsert_meta_dict(conn: Any, email: str, meta: dict[str, Any]) -> None:
    email = email.strip()
    if not email:
        return
    verify = meta.get("verify")
    conn.execute(
        """INSERT INTO account_meta(email, note, tags_json, verify_json, combo_dual_meta, updated_at)
           VALUES (?,?,?,?,?,?)
           ON CONFLICT(email) DO UPDATE SET
             note=excluded.note,
             tags_json=excluded.tags_json,
             verify_json=excluded.verify_json,
             combo_dual_meta=excluded.combo_dual_meta,
             updated_at=excluded.updated_at
        """,
        (
            email,
            meta.get("note") or "",
            json.dumps(meta.get("tags") or [], ensure_ascii=False),
            json.dumps(verify, ensure_ascii=False) if verify is not None else "",
            meta.get("combo_dual") or "",
            meta.get("updated_at") or _now(),
        ),
    )


def save_register_result(
    result: RegisterResult,
    output_dir: str = "accounts",
    *,
    batch_id: str = "",
    batch_no: Optional[int] = None,
    batch_label: str = "",
) -> str:
    del output_dir  # 统一 SQLite，目录参数保留兼容
    db.ensure_initialized()
    result = enrich_register_result(result)
    has_dual = bool(result.login_refresh_token)
    has_recovery = bool(result.recovery_email and result.recovery_password)
    auth_status = "ok" if result.refresh_token else "missing_token"
    now = _now()
    row = {
        "email": result.email,
        "password": result.password,
        "redirect_url": result.redirect_url,
        "refresh_token": result.refresh_token or "",
        "client_id": result.client_id or "",
        "login_client_id": result.login_client_id or "",
        "login_refresh_token": result.login_refresh_token or "",
        "recovery_email": result.recovery_email or "",
        "recovery_password": result.recovery_password or "",
        "combo": result.to_combo(),
        "combo_dual": result.to_combo(dual=True) if has_dual else "",
        "combo_recovery": result.to_combo(recovery=True) if has_recovery else "",
        "dual_requested": bool(result.extra.get("dual_requested")),
        "dual_ok": bool(result.extra.get("dual_ok")),
        "login_status": result.extra.get("login_status", ""),
        "login_fail_reason": result.extra.get("login_fail_reason", ""),
        "auth_status": auth_status,
        "proofs_method": result.extra.get("proofs_method", ""),
        "proofs_satisfied": result.extra.get("proofs_satisfied", ""),
        "success": result.success,
        "error": result.error or "",
        "extra": result.extra,
        "created_at": now,
        "updated_at": now,
        "batch_id": batch_id or "",
        "batch_no": batch_no,
        "batch_label": batch_label or (f"B{batch_no}" if batch_no else ""),
        "legacy_source": "register",
    }
    with _lock:
        conn = db.connect()
        try:
            upsert_account_dict(conn, row)
            # 新号默认打 incubating 标签（时长由 OUTLOOK_INCUBATION_HOURS 决定）
            meta_row = conn.execute(
                "SELECT tags_json FROM account_meta WHERE email=? COLLATE NOCASE",
                (result.email,),
            ).fetchone()
            tags: list[Any] = []
            if meta_row and meta_row["tags_json"]:
                try:
                    tags = json.loads(meta_row["tags_json"] or "[]")
                except Exception:  # noqa: BLE001
                    tags = []
            if not isinstance(tags, list):
                tags = []
            if INCUBATING_TAG not in tags:
                tags.append(INCUBATING_TAG)
            upsert_meta_dict(conn, result.email, {"tags": tags, "updated_at": now})
            conn.commit()
        finally:
            conn.close()
    logger.info("账号已保存（孵化中）: %s", result.email)
    return f"sqlite:{result.email}"


def get_account(email: str) -> Optional[dict[str, Any]]:
    db.ensure_initialized()
    email = email.strip()
    conn = db.connect()
    try:
        row = conn.execute("SELECT * FROM accounts WHERE email=? COLLATE NOCASE", (email,)).fetchone()
        if not row:
            return None
        acc = _row_to_dict(row)
        meta = conn.execute("SELECT * FROM account_meta WHERE email=? COLLATE NOCASE", (email,)).fetchone()
        if meta:
            acc["note"] = meta["note"] or ""
            try:
                acc["tags"] = json.loads(meta["tags_json"] or "[]")
            except Exception:  # noqa: BLE001
                acc["tags"] = []
            if meta["verify_json"]:
                try:
                    acc["verify"] = json.loads(meta["verify_json"])
                except Exception:  # noqa: BLE001
                    acc["verify"] = None
            if not acc.get("combo_dual") and meta["combo_dual_meta"]:
                acc["combo_dual"] = meta["combo_dual_meta"]
        else:
            acc.setdefault("tags", [])
            acc.setdefault("note", "")
        enrich_lifecycle_fields(acc)
        return acc
    finally:
        conn.close()


def patch_account(email: str, patch: dict[str, Any]) -> None:
    acc = get_account(email) or {"email": email}
    acc.update(patch)
    acc["updated_at"] = patch.get("updated_at") or _now()
    with _lock:
        conn = db.connect()
        try:
            upsert_account_dict(conn, acc)
            conn.commit()
        finally:
            conn.close()


def update_meta(email: str, patch: dict[str, Any]) -> None:
    db.ensure_initialized()
    email = email.strip()
    with _lock:
        conn = db.connect()
        try:
            row = conn.execute(
                "SELECT * FROM account_meta WHERE email=? COLLATE NOCASE", (email,)
            ).fetchone()
            cur: dict[str, Any] = {}
            if row:
                cur = {
                    "note": row["note"],
                    "tags": json.loads(row["tags_json"] or "[]"),
                    "combo_dual": row["combo_dual_meta"],
                }
                if row["verify_json"]:
                    try:
                        cur["verify"] = json.loads(row["verify_json"])
                    except Exception:  # noqa: BLE001
                        pass
            cur.update(patch)
            cur["updated_at"] = _now()
            upsert_meta_dict(conn, email, cur)
            conn.commit()
        finally:
            conn.close()


def cache_verify(email: str, res: dict[str, Any]) -> None:
    if res.get("transient") or res.get("unable"):
        return
    update_meta(email, {
        "verify": {
            "ok": res.get("ok"),
            "usable": res.get("usable", []),
            "graph": res.get("graph", {}).get("ok") if isinstance(res.get("graph"), dict) else res.get("graph"),
            "outlook_rest": res.get("outlook_rest", {}).get("ok")
            if isinstance(res.get("outlook_rest"), dict)
            else res.get("outlook_rest"),
            "imap": res.get("imap", {}).get("ok") if isinstance(res.get("imap"), dict) else res.get("imap"),
            "granted_scope": res.get("granted_scope", ""),
            "checked_at": _now(),
        }
    })
    patch: dict[str, Any] = {"updated_at": _now()}
    if res.get("ok"):
        patch["last_alive_at"] = _now()
    patch_account(email, patch)


def delete_accounts(emails: list[str]) -> int:
    emails = [e.strip() for e in emails if (e or "").strip()]
    if not emails:
        return 0
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            n = 0
            for e in emails:
                cur = conn.execute("DELETE FROM accounts WHERE email=? COLLATE NOCASE", (e,))
                n += cur.rowcount
            conn.commit()
            return n
        finally:
            conn.close()


def rescue_counts() -> dict[str, int]:
    db.ensure_initialized()
    conn = db.connect()
    try:
        out: dict[str, int] = {}
        for r in conn.execute("SELECT email, rescue_count FROM accounts"):
            out[r["email"].lower()] = int(r["rescue_count"] or 0)
        for r in conn.execute("SELECT email, COUNT(*) AS c FROM rescue_events GROUP BY email"):
            out[r["email"].lower()] = max(out.get(r["email"].lower(), 0), int(r["c"]))
        return out
    finally:
        conn.close()


def write_rescue_outcome(email: str, data: dict[str, Any], out: dict[str, Any]) -> str:
    """救援结果写回 SQLite（替代 JSON writeback）。"""
    email = (email or data.get("email") or "").strip()
    if not email:
        return ""
    data = dict(data)
    data["email"] = email
    n = int(data.get("rescue_count") or 0) + 1
    data["rescue_count"] = n
    data["last_rescue_at"] = _now()
    data["last_rescue_ok"] = bool(out.get("ok"))
    reason = str(out.get("reason") or out.get("message") or "").strip()
    data["last_rescue_reason"] = "" if out.get("ok") else reason[:240]

    rt = out.get("refresh_token", "") if out.get("ok") else ""
    if rt:
        pwd = data.get("password", "")
        cid = data.get("client_id") or ""
        data["refresh_token"] = rt
        data["combo"] = "----".join([email, pwd, cid, rt])
        data["rescued_at"] = _now()
        data["last_alive_at"] = _now()
        data["rescued_scope"] = out.get("scope") or ""
    rec_e = data.get("recovery_email") or ""
    rec_p = data.get("recovery_password") or ""
    if rec_e:
        rt_seg = rt or data.get("refresh_token") or ""
        data["combo_recovery"] = "----".join([email, data.get("password", ""), data.get("client_id", ""), rt_seg, rec_e, rec_p])

    data["updated_at"] = _now()
    with _lock:
        conn = db.connect()
        try:
            upsert_account_dict(conn, data)
            conn.execute(
                "INSERT INTO rescue_events(email, ok, reason, created_at) VALUES (?,?,?,?)",
                (email.lower(), 1 if out.get("ok") else 0, reason[:240], _now()),
            )
            conn.commit()
        finally:
            conn.close()
    return f"sqlite:{email}"


def save_job(record: dict[str, Any]) -> None:
    db.ensure_initialized()
    with _lock:
        conn = db.connect()
        try:
            conn.execute(
                """INSERT OR REPLACE INTO register_jobs(
                    id, batch_no, batch_label, created_at, status, count, concurrency,
                    token_mode, dry_run, ok_count, fail_count, emails_json, params_json
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    record.get("id") or "",
                    record.get("batch_no"),
                    record.get("batch_label") or "",
                    record.get("created_at") or "",
                    record.get("status") or "",
                    int(record.get("count") or 0),
                    int(record.get("concurrency") or 1),
                    record.get("token_mode") or "",
                    1 if record.get("dry_run") else 0,
                    int(record.get("ok_count") or 0),
                    int(record.get("fail_count") or 0),
                    json.dumps(record.get("emails") or [], ensure_ascii=False),
                    json.dumps(record.get("params") or {}, ensure_ascii=False),
                ),
            )
            conn.commit()
        finally:
            conn.close()


def list_jobs(limit: int = 200) -> list[dict[str, Any]]:
    db.ensure_initialized()
    conn = db.connect()
    try:
        rows = conn.execute(
            "SELECT * FROM register_jobs ORDER BY created_at DESC LIMIT ?",
            (max(1, limit),),
        ).fetchall()
        out: list[dict[str, Any]] = []
        for r in rows:
            rec = dict(r)
            try:
                rec["emails"] = json.loads(rec.pop("emails_json") or "[]")
            except Exception:  # noqa: BLE001
                rec["emails"] = []
            rec.pop("params_json", None)
            out.append(rec)
        return out
    finally:
        conn.close()


def batch_index() -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    for rec in list_jobs():
        info = {
            "batch_id": rec.get("id") or "",
            "batch_no": rec.get("batch_no"),
            "batch_label": rec.get("batch_label") or "",
        }
        for em in rec.get("emails") or []:
            if em:
                out[str(em)] = info
    return out


# ---------------------------------------------------------------------------
# 账号池分页查询（SQL 侧过滤 + LIMIT/OFFSET，避免近万行整表搬进内存/网络）。
#
# 列表行只带渲染所需的瘦身字段：refresh_token / combo / combo_dual /
# login_refresh_token 只在这里用来算 has_token / token_tail / login_token，
# 不进最终返回值；完整明文只在 get_account()（单号详情弹窗）里给。
# ---------------------------------------------------------------------------

_VIEW_SQL: dict[str, str] = {
    "without": "COALESCE(a.refresh_token,'')=''",
    "dual": "(COALESCE(a.combo_dual,'')!='' OR COALESCE(a.login_refresh_token,'')!='' OR COALESCE(m.combo_dual_meta,'')!='')",
    "untested": "(m.verify_json IS NULL OR m.verify_json='')",
    "usable": (
        "(m.verify_json IS NOT NULL AND m.verify_json!='' "
        "AND (json_extract(m.verify_json,'$.ok')=1 OR json_extract(m.verify_json,'$.graph')=1))"
    ),
    "dead": (
        "(m.verify_json IS NOT NULL AND m.verify_json!='' "
        "AND COALESCE(json_extract(m.verify_json,'$.ok'),0)=0 "
        "AND COALESCE(json_extract(m.verify_json,'$.graph'),0)=0)"
    ),
}

# 批次列没填时按 batch_no 拼 "B{no}"，跟前端历史展示口径（accountBatchLabel）一致。
_EFFECTIVE_BATCH_SQL = (
    "(CASE WHEN COALESCE(a.batch_label,'')!='' THEN a.batch_label "
    "WHEN a.batch_no IS NOT NULL THEN 'B'||a.batch_no ELSE '' END)"
)

_LIST_JOIN = "FROM accounts a LEFT JOIN account_meta m ON a.email = m.email COLLATE NOCASE"

_LIST_COLUMNS = (
    "a.email", "a.password", "a.recovery_email",
    "a.refresh_token", "a.combo_dual", "a.login_refresh_token",
    "a.batch_id", "a.batch_no", "a.batch_label", "a.rescue_count",
    "a.last_rescue_at", "a.last_rescue_ok", "a.last_rescue_reason",
    "a.created_at", "a.updated_at", "a.last_alive_at", "a.legacy_source",
    "m.note", "m.tags_json", "m.verify_json", "m.combo_dual_meta",
)


def _accounts_query_where(filters: dict[str, Any]) -> tuple[str, list[Any]]:
    conds: list[str] = []
    params: list[Any] = []
    view = (filters.get("view") or "all").strip().lower()
    if view in _VIEW_SQL:
        conds.append(_VIEW_SQL[view])
    batch = (filters.get("batch") or "all").strip()
    if batch == "none":
        conds.append(f"{_EFFECTIVE_BATCH_SQL}=''")
    elif batch and batch.lower() != "all":
        conds.append(f"{_EFFECTIVE_BATCH_SQL}=?")
        params.append(batch)
    q = (filters.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        conds.append(
            "(a.email LIKE ? OR COALESCE(a.recovery_email,'') LIKE ? "
            "OR COALESCE(m.note,'') LIKE ? OR COALESCE(a.batch_label,'') LIKE ? "
            "OR COALESCE(m.tags_json,'') LIKE ?)"
        )
        params.extend([like, like, like, like, like])
    return (" AND ".join(conds) if conds else "1=1"), params


def count_accounts(filters: Optional[dict[str, Any]] = None) -> int:
    """按 q / batch / view 过滤后的总数，用于分页页码。"""
    db.ensure_initialized()
    conn = db.connect()
    try:
        where, params = _accounts_query_where(filters or {})
        row = conn.execute(f"SELECT COUNT(*) AS c {_LIST_JOIN} WHERE {where}", params).fetchone()
        return int(row["c"]) if row else 0
    finally:
        conn.close()


def rescue_count_for_email(email: str) -> int:
    """单号详情用：只查这一个邮箱的 rescue_events，不像 rescue_counts() 那样扫全表。"""
    email = (email or "").strip()
    if not email:
        return 0
    db.ensure_initialized()
    conn = db.connect()
    try:
        row = conn.execute(
            "SELECT COUNT(*) AS c FROM rescue_events WHERE email=?", (email.lower(),)
        ).fetchone()
        return int(row["c"] or 0) if row else 0
    finally:
        conn.close()


def _rescue_counts_for(conn: Any, emails: list[str]) -> dict[str, int]:
    """只对当前这一页的邮箱查 rescue_events，不像 rescue_counts() 那样扫全表。"""
    if not emails:
        return {}
    placeholders = ",".join("?" for _ in emails)
    rows = conn.execute(
        f"SELECT email, COUNT(*) AS c FROM rescue_events WHERE email IN ({placeholders}) GROUP BY email",
        [e.lower() for e in emails],
    ).fetchall()
    return {r["email"].lower(): int(r["c"] or 0) for r in rows}


def list_accounts_page(
    limit: int = 50,
    offset: int = 0,
    q: str = "",
    batch: str = "all",
    view: str = "all",
) -> list[dict[str, Any]]:
    """账号池分页列表：SQL 侧过滤 + LIMIT/OFFSET + 只选必要列。

    排序与旧版 list_accounts() 的前端排序口径一致：可读信的排前面，同组内按
    created_at 倒序。
    """
    db.ensure_initialized()
    conn = db.connect()
    try:
        where, params = _accounts_query_where({"q": q, "batch": batch, "view": view})
        order_sql = f"CASE WHEN {_VIEW_SQL['usable']} THEN 1 ELSE 0 END DESC, a.created_at DESC"
        sql = (
            f"SELECT {', '.join(_LIST_COLUMNS)} {_LIST_JOIN} WHERE {where} "
            f"ORDER BY {order_sql} LIMIT ? OFFSET ?"
        )
        rows = conn.execute(sql, params + [max(1, int(limit)), max(0, int(offset))]).fetchall()
        emails = [r["email"] for r in rows]
        rescue_extra = _rescue_counts_for(conn, emails)
        out: list[dict[str, Any]] = []
        for r in rows:
            d = dict(r)
            email = d.get("email") or ""
            refresh_token = d.pop("refresh_token", "") or ""
            combo_dual = d.pop("combo_dual", "") or ""
            login_refresh_token = d.pop("login_refresh_token", "") or ""
            combo_dual_meta = d.pop("combo_dual_meta", "") or ""
            d["has_token"] = bool(refresh_token)
            d["token_tail"] = refresh_token[-6:] if refresh_token else ""
            d["login_token"] = bool(combo_dual or login_refresh_token or combo_dual_meta)
            d["has_recovery"] = bool(d.get("recovery_email"))
            try:
                d["tags"] = json.loads(d.pop("tags_json", None) or "[]")
            except Exception:  # noqa: BLE001
                d["tags"] = []
            verify_json = d.pop("verify_json", None)
            try:
                d["verify"] = json.loads(verify_json) if verify_json else None
            except Exception:  # noqa: BLE001
                d["verify"] = None
            d["note"] = d.get("note") or ""
            d["source"] = d.get("legacy_source") or "sqlite"
            rescue_n = rescue_extra.get(email.lower(), 0)
            if rescue_n:
                d["rescue_count"] = max(int(d.get("rescue_count") or 0), rescue_n)
            enrich_lifecycle_fields(d)
            out.append(d)
        return out
    finally:
        conn.close()


def overview_stats() -> dict[str, Any]:
    """概览统计：全走 SQL 聚合，不把整表读进 Python。"""
    db.ensure_initialized()
    conn = db.connect()
    try:
        total = int(conn.execute(f"SELECT COUNT(*) AS c {_LIST_JOIN}").fetchone()["c"])
        with_token = int(
            conn.execute(
                f"SELECT COUNT(*) AS c {_LIST_JOIN} WHERE COALESCE(a.refresh_token,'')!=''"
            ).fetchone()["c"]
        )
        usable = int(
            conn.execute(f"SELECT COUNT(*) AS c {_LIST_JOIN} WHERE {_VIEW_SQL['usable']}").fetchone()["c"]
        )
        # 无批次的旧号失败测活不计入失活，避免数字膨胀（与旧版 _compute_stats 口径一致）
        dead = int(
            conn.execute(
                f"SELECT COUNT(*) AS c {_LIST_JOIN} WHERE {_VIEW_SQL['dead']} "
                "AND (COALESCE(a.batch_label,'')!='' OR a.batch_no IS NOT NULL)"
            ).fetchone()["c"]
        )
        untested = int(
            conn.execute(f"SELECT COUNT(*) AS c {_LIST_JOIN} WHERE {_VIEW_SQL['untested']}").fetchone()["c"]
        )
        today = _now()[:10]
        today_new = int(
            conn.execute(
                "SELECT COUNT(*) AS c FROM accounts WHERE created_at LIKE ?", (today + "%",)
            ).fetchone()["c"]
        )
        batch_rows = conn.execute(
            f"SELECT {_EFFECTIVE_BATCH_SQL} AS label, COUNT(*) AS c FROM accounts a "
            "GROUP BY label HAVING label!='' ORDER BY label DESC"
        ).fetchall()
        return {
            "total": total,
            "with_token": with_token,
            "usable": usable,
            "dead": dead,
            "untested": untested,
            "today_new": today_new,
            "batches": [{"label": r["label"], "count": int(r["c"])} for r in batch_rows],
        }
    finally:
        conn.close()


def list_accounts() -> list[dict[str, Any]]:
    """Web 账号池列表（SQLite 为唯一数据源）。"""
    db.ensure_initialized()
    conn = db.connect()
    try:
        accounts = [_row_to_dict(r) for r in conn.execute("SELECT * FROM accounts")]
        meta_map: dict[str, dict[str, Any]] = {}
        for m in conn.execute("SELECT * FROM account_meta"):
            email = m["email"]
            verify = None
            if m["verify_json"]:
                try:
                    verify = json.loads(m["verify_json"])
                except Exception:  # noqa: BLE001
                    verify = None
            meta_map[email] = {
                "note": m["note"] or "",
                "tags": json.loads(m["tags_json"] or "[]"),
                "verify": verify,
                "combo_dual": m["combo_dual_meta"] or "",
            }
        rescue_log = rescue_counts()
        batch_map = batch_index()
        for row in accounts:
            email = row["email"]
            m = meta_map.get(email, {})
            row["note"] = m.get("note", "")
            row["tags"] = m.get("tags", [])
            row["verify"] = m.get("verify")
            if not row.get("combo_dual") and m.get("combo_dual"):
                row["combo_dual"] = m["combo_dual"]
                row["login_token"] = True
            hit = batch_map.get(email) or {}
            if hit.get("batch_id") and not row.get("batch_id"):
                row["batch_id"] = hit["batch_id"]
            if hit.get("batch_no") and not row.get("batch_no"):
                row["batch_no"] = hit["batch_no"]
            if hit.get("batch_label") and not row.get("batch_label"):
                row["batch_label"] = hit["batch_label"]
            log_n = rescue_log.get(email.lower(), 0)
            if log_n:
                row["rescue_count"] = max(int(row.get("rescue_count") or 0), log_n)
            if not row.get("combo"):
                row["combo"] = "----".join([
                    email,
                    row.get("password", ""),
                    row.get("client_id", ""),
                    row.get("refresh_token", ""),
                ])
            enrich_lifecycle_fields(row)
        accounts.sort(
            key=lambda x: (
                1 if (x.get("verify") or {}).get("ok") or (x.get("verify") or {}).get("graph") else 0,
                x.get("created_at") or "",
            ),
            reverse=True,
        )
        return accounts
    finally:
        conn.close()
