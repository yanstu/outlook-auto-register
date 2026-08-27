"""Outlook 纯 API 注册工具的 Web 后端（FastAPI）。

设计原则：
- 只 import 调用现有 `outlook_api_reg` 的能力（register_one / [可选]register_batch /
  save_account / models / proxy_utils / mail_reader / graph_mail / enable_imap），
  绝不修改核心协议流程文件（另一个 worker 正在改引擎）。
- 注册是阻塞式网络任务：放后台线程 + 线程池并发执行，通过 SSE 实时推进度。
- 邮件令牌模式 graph/outlook_rest/imap/dual（由引擎 _MAIL_SCOPE_BY_MODE 决定实际可用）；
  默认 graph 走 Graph API 读信，绕开新号 IMAP 开关限制。
- 批量注册走引擎 register_batch_iter（生成器驱动 SSE，引擎内部已 save_account），
  import 失败时线程池兜底。
- 保活接 scripts.keepalive.keepalive_one（refresh→access→GET /me+列信→轮换回写）；
  「开启 IMAP」仍为占位，如实反馈，绝不假装成功。

启动：
  cd outlook-api-register
  .venv/bin/uvicorn webapp.server:app --host 0.0.0.0 --port 8890
"""
from __future__ import annotations

import json
import logging
import os
import queue
import re
import sys
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import (
    HTMLResponse,
    JSONResponse,
    PlainTextResponse,
    StreamingResponse,
)
from pydantic import BaseModel

# 现有能力（只读 import 调用，不修改核心协议文件）
from outlook_api_reg.account_persist import merge_account_row
from outlook_api_reg import account_store
from outlook_api_reg import database as app_db
from outlook_api_reg.register import register_one, save_account
from outlook_api_reg.models import RegisterResult
from outlook_api_reg import mail_reader
from outlook_api_reg import graph_mail, enable_imap, post_register
from outlook_api_reg import constants as reg_constants
from outlook_api_reg import cf_domain_mail
from outlook_api_reg import coolhs_mail
from outlook_api_reg import proxy_pool
from outlook_api_reg import external_recovery_pool as ext_recovery_pool
from outlook_api_reg import register as reg_module

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
ACCOUNTS_DIR = PROJECT_DIR / "accounts"
STATIC_DIR = BASE_DIR / "static"
META_FILE = ACCOUNTS_DIR / "webapp_meta.json"
JOBS_FILE = ACCOUNTS_DIR / "webapp_jobs.json"
COMBO_FILE = ACCOUNTS_DIR / "accounts.txt"
DUAL_FILE = ACCOUNTS_DIR / "accounts_dual.txt"
RECOVERY_FILE = ACCOUNTS_DIR / "accounts_recovery.txt"

# 保活脚本位于 scripts/（非包）：把仓库根加入 sys.path 后 import keepalive_one
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))
try:
    from scripts.keepalive import keepalive_one
    KEEPALIVE_READY = True
except Exception:  # noqa: BLE001
    keepalive_one = None
    KEEPALIVE_READY = False

try:
    from scripts.rescue_login import count_rescues_from_log, rescue_and_persist, _proxy_raw as rescue_proxy_raw
    RESCUE_READY = True
except Exception:  # noqa: BLE001
    count_rescues_from_log = None  # type: ignore[misc, assignment]
    rescue_and_persist = None  # type: ignore[misc, assignment]
    rescue_proxy_raw = None  # type: ignore[misc, assignment]
    RESCUE_READY = False

logger = logging.getLogger(__name__)
MAIL_CLIENT_ID = os.environ.get("MAIL_CLIENT_ID", "9e5f94bc-e8a4-4e73-b8be-63364c29d753")

# 邮件令牌模式（引擎里的 MAIL_SCOPE_BY_MODE 决定实际可用者）
_scope_map = getattr(reg_constants, "MAIL_SCOPE_BY_MODE", None) or getattr(
    reg_constants, "_MAIL_SCOPE_BY_MODE", {}
)
_ENGINE_MODES = list(_scope_map.keys()) or [
    "graph",
    "outlook_rest",
    "imap",
]
DEFAULT_TOKEN_MODE = getattr(reg_constants, "MAIL_TOKEN_MODE", "graph")
DUAL_READY = "dual" in _scope_map
# 引擎高级模式仍暴露给 API；页面「产出格式」只用 graph / login_exe
TOKEN_MODES = _ENGINE_MODES + (["dual"] if "dual" not in _ENGINE_MODES else [])
PRODUCT_MODES = [
    {
        "id": "graph",
        "label": "Graph 四段式",
        "export": "graph",
        "hint": "",
    },
    {
        "id": "graph_recovery",
        "label": "Graph 六段式（推荐）",
        "export": "recovery",
        "hint": "",
    },
]
EXPORT_FORMATS = ["graph", "recovery", "dual"]
# 引擎批量注册生成器（特性探测；import 失败则线程池兜底）
try:
    from outlook_api_reg.batch import register_batch_iter
    BATCH_READY = True
except Exception:  # noqa: BLE001
    register_batch_iter = None
    BATCH_READY = False

app = FastAPI(title="Outlook API 注册控制台", version="2.0.0")

_save_lock = threading.Lock()
_meta_lock = threading.Lock()


def _proxy_url(raw: Optional[str]) -> str:
    if not raw:
        return ""
    try:
        from outlook_api_reg.proxy_utils import parse_proxy

        cfg = parse_proxy(raw)
        return cfg.url if cfg else ""
    except Exception:  # noqa: BLE001
        return ""


def _build_register_proxy_plan(p: dict[str, Any], count: int) -> tuple[list[str], dict[str, Any]]:
    """注册任务：代理池（SQLite）优先，备用代理仅来自 Web 表单。"""
    if p.get("use_proxy_pool"):
        plan, meta = proxy_pool.plan_for_batch(count)
        if len(plan) < count:
            fallback = (p.get("proxy") or "").strip()
            if fallback:
                from outlook_api_reg.proxy_utils import expand_proxy_unique

                extra = expand_proxy_unique(fallback, count - len(plan))
                plan = plan + list(extra)
                meta["fallback_used"] = True
        if not plan:
            raise ValueError("代理池无可用条目，请先在「代理池」页添加代理，或在注册页填写备用代理（会自动写入数据库）。")
        return plan, meta
    proxy = (p.get("proxy") or "").strip()
    if not proxy:
        raise ValueError("请填写代理或启用「使用代理池」。")
    from outlook_api_reg.batch import _plan_proxies

    return [x or "" for x in _plan_proxies(proxy, count)], {"source": "manual"}


def _after_register_proxy(
    email: str,
    assignments: list[dict[str, Any]],
    index: int,
    *,
    success: bool,
    reg_country: str = "US",
    error: str = "",
) -> None:
    if index >= len(assignments):
        return
    a = assignments[index]
    pid = a.get("proxy_id") or ""
    resolved = a.get("resolved") or ""
    if success and email and pid and resolved:
        proxy_pool.bind_account(email, pid, resolved, purpose="register")
    if pid:
        proxy_pool.record_result(
            pid,
            success=success,
            reg_country=reg_country,
            purpose="register",
            email=email if success else "",
            error=error if not success else "",
        )


def _apply_token_mode(mode: str) -> str:
    """按选择切换邮件令牌 scope，返回产出格式 mode（graph_recovery 不降级为 graph）。"""
    normalize = getattr(reg_constants, "normalize_token_mode", lambda m: (m or "graph").strip().lower())
    mode = normalize(mode)
    scope_map = getattr(reg_constants, "MAIL_SCOPE_BY_MODE", None) or getattr(
        reg_constants, "_MAIL_SCOPE_BY_MODE", {}
    )
    is_recovery = getattr(reg_constants, "is_recovery_mode", lambda m: m in ("recovery", "login_exe"))
    is_graph_recovery = getattr(reg_constants, "is_graph_recovery_mode", lambda m: m == "graph_recovery")

    dual = False
    if is_graph_recovery(mode):
        effective = "graph_recovery"
        scope = scope_map.get("graph")
    elif mode == "dual" and not DUAL_READY:
        effective = "graph"
        scope = scope_map.get("graph")
    else:
        scope = scope_map.get(mode)
        effective = mode
        if not scope:
            effective = "graph"
            scope = scope_map.get("graph")
        dual = effective == "dual"
        if is_recovery(effective):
            dual = False
            scope = scope or scope_map.get("login_exe") or scope_map.get("imap")

    reg_constants.DUAL_TOKEN = dual
    post_register.DUAL_TOKEN = dual
    reg_module.DUAL_TOKEN = dual
    oauth_mode = "graph" if is_graph_recovery(mode) else effective
    if hasattr(reg_constants, "MAIL_TOKEN_MODE"):
        reg_constants.MAIL_TOKEN_MODE = oauth_mode
    if scope:
        os.environ["OUTLOOK_MAIL_TOKEN_MODE"] = oauth_mode
        post_register.MAIL_SCOPE = scope
        reg_constants.MAIL_SCOPE = scope
    return effective


def _job_combo(result: RegisterResult, mode: str) -> str:
    if hasattr(result, "product_combo"):
        return result.product_combo(mode)
    return result.to_combo()


# ---------------------------------------------------------------------------
# 账号元数据（备注 / 标签 / 测活缓存）——存 accounts/webapp_meta.json
# ---------------------------------------------------------------------------


def _load_meta() -> dict[str, Any]:
    """遗留兼容：元数据已迁入 SQLite account_meta。"""
    return {}


def _save_meta(meta: dict[str, Any]) -> None:
    del meta


def _update_meta(email: str, patch: dict[str, Any]) -> None:
    account_store.update_meta(email, patch)


_FN_TS = re.compile(r"_(\d{8})_(\d{6})\.json$")
_BATCH_LABEL_SAFE = re.compile(r"[^\w\-.@+]+")


def _domain_slug(domain: str) -> str:
    d = (domain or "@outlook.com").strip().lstrip("@").split(".")[0]
    return (d or "outlook").lower()


def _token_mode_slug(mode: str) -> str:
    m = (mode or "graph").strip().lower()
    return {
        "graph": "G4",
        "graph_recovery": "G6",
        "recovery": "IMAP6",
        "login_exe": "IMAP6",
        "dual": "DUAL",
        "outlook_rest": "REST",
        "imap": "IMAP",
    }.get(m, m[:8].upper())


def _sanitize_batch_label(raw: str) -> str:
    s = _BATCH_LABEL_SAFE.sub("-", (raw or "").strip())
    s = re.sub(r"-+", "-", s).strip("-")
    return s[:48]


def _existing_batch_labels(*, jobs: Optional[dict[str, Any]] = None) -> set[str]:
    labels: set[str] = set()
    for rec in _load_jobs_store():
        lb = _sanitize_batch_label(str(rec.get("batch_label") or ""))
        if lb:
            labels.add(lb)
    src = jobs if jobs is not None else _jobs
    if jobs is None:
        with _jobs_lock:
            for j in src.values():
                lb = _sanitize_batch_label(str(getattr(j, "batch_label", "") or ""))
                if lb:
                    labels.add(lb)
    else:
        for j in src.values():
            lb = _sanitize_batch_label(str(getattr(j, "batch_label", "") or ""))
            if lb:
                labels.add(lb)
    return labels


def _make_batch_label(params: dict[str, Any], batch_no: int, *, jobs: Optional[dict[str, Any]] = None) -> str:
    """生成有业务含义的批次名；用户可传 batch_label 覆盖。"""
    custom = _sanitize_batch_label(str(params.get("batch_label") or ""))
    if custom:
        return custom
    country = (params.get("country") or "US").strip().upper()[:6]
    dom = _domain_slug(str(params.get("domain") or "@outlook.com"))
    count = max(1, int(params.get("count") or 1))
    mode = _token_mode_slug(str(params.get("token_mode") or "graph"))
    prefix = _sanitize_batch_label(str(params.get("prefix") or ""))
    date_part = datetime.now().strftime("%m%d")
    parts = [date_part, country, dom]
    if prefix:
        parts.append(prefix[:12])
    parts.extend([f"{count}x", mode])
    label = "-".join(p for p in parts if p)
    if label in _existing_batch_labels(jobs=jobs):
        label = f"{label}-#{batch_no}"
    return label[:48]


def _parse_dt(raw: str) -> Optional[datetime]:
    s = (raw or "").strip()
    if not s:
        return None
    s = s.replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


def _dt_iso(dt: Optional[datetime]) -> str:
    return dt.isoformat() if dt else ""


def _infer_created_at(fp: Path, data: dict[str, Any]) -> str:
    existing = _parse_dt(str(data.get("created_at") or ""))
    if existing:
        return _dt_iso(existing)
    m = _FN_TS.search(fp.name)
    if m:
        try:
            return datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S").isoformat()
        except ValueError:
            pass
    try:
        return datetime.fromtimestamp(fp.stat().st_mtime).isoformat()
    except OSError:
        return ""


def _best_updated_at(data: dict[str, Any], meta_entry: dict[str, Any], fp: Optional[Path] = None) -> str:
    cands: list[datetime] = []
    for raw in (
        data.get("updated_at"),
        data.get("rescued_at"),
        data.get("last_alive_at"),
        (meta_entry or {}).get("updated_at"),
        ((meta_entry or {}).get("verify") or {}).get("checked_at"),
    ):
        dt = _parse_dt(str(raw or ""))
        if dt:
            cands.append(dt.replace(tzinfo=None) if dt.tzinfo else dt)
    if fp is not None:
        try:
            cands.append(datetime.fromtimestamp(fp.stat().st_mtime))
        except OSError:
            pass
    if not cands:
        created = _parse_dt(str(data.get("created_at") or ""))
        return _dt_iso(created.replace(tzinfo=None) if created and created.tzinfo else created)
    return max(cands).isoformat()


def _best_last_alive(data: dict[str, Any], meta_entry: dict[str, Any]) -> str:
    verify = (meta_entry or {}).get("verify") or {}
    if verify.get("checked_at"):
        return str(verify["checked_at"])
    for key in ("last_alive_at", "rescued_at"):
        if data.get(key):
            return str(data[key])
    return ""


def _survival_end_at(row: dict[str, Any], meta_entry: dict[str, Any]) -> Optional[datetime]:
    """存活统计终点：首次测活/保活确认（活或死）的时刻，不用「当前时间」。"""
    verify = (meta_entry or {}).get("verify")
    if isinstance(verify, dict) and verify.get("checked_at"):
        return _parse_dt(str(verify["checked_at"]))
    return None


def _compute_alive_seconds(created_raw: str, end_dt: Optional[datetime]) -> Optional[int]:
    created = _parse_dt(created_raw)
    if not created or not end_dt:
        return None
    created_naive = created.replace(tzinfo=None) if created.tzinfo else created
    end_naive = end_dt.replace(tzinfo=None) if end_dt.tzinfo else end_dt
    if end_naive < created_naive:
        return None
    return max(0, int((end_naive - created_naive).total_seconds()))


def _patch_account_json(email: str, patch: dict[str, Any]) -> None:
    account_store.patch_account(email, patch)


# ---------------------------------------------------------------------------
# 任务 / 日志基础设施
# ---------------------------------------------------------------------------

_jobs: dict[str, "Job"] = {}
_jobs_lock = threading.Lock()
_jobs_file_lock = threading.Lock()
_thread_job: dict[int, str] = {}  # 线程 ident -> job_id，日志归属
_active_batch_job: Optional[str] = None  # 引擎批量运行时：其内部线程池日志归属到此任务


class Job:
    def __init__(self, job_id: str, params: dict[str, Any]):
        self.id = job_id
        self.params = params
        self.status = "running"
        self.created_at = datetime.now().isoformat()
        self.count = int(params.get("count", 1))
        self.concurrency = int(params.get("concurrency", 1) or 1)
        self.batch_no = int(params.get("batch_no") or 0)
        self.batch_label = str(params.get("batch_label") or (f"B{self.batch_no}" if self.batch_no else ""))
        self.accounts: list[dict[str, Any]] = [
            {
                "index": i + 1,
                "status": "等待中",
                "email": "",
                "password": "",
                "client_id": "",
                "refresh_token": "",
                "combo": "",
                "combo_dual": "",
                "recovery_email": "",
                "recovery_password": "",
                "login_token": False,
                "error": "",
                "saved_path": "",
            }
            for i in range(self.count)
        ]
        self.logs: list[dict[str, Any]] = []
        self.batch_summary: Optional[dict[str, Any]] = None
        self._queue: "queue.Queue[dict[str, Any]]" = queue.Queue()
        self._lock = threading.Lock()

    def emit(self, event: dict[str, Any]) -> None:
        self._queue.put(event)

    def push_log(self, level: str, msg: str) -> None:
        rec = {"ts": datetime.now().strftime("%H:%M:%S"), "level": level, "msg": msg}
        with self._lock:
            self.logs.append(rec)
            if len(self.logs) > 2000:
                self.logs = self.logs[-2000:]
        self.emit({"type": "log", **rec})

    def update_account(self, index0: int, **fields: Any) -> None:
        with self._lock:
            self.accounts[index0].update(fields)
            snapshot = dict(self.accounts[index0])
        self.emit({"type": "account", "account": snapshot})

    def counts(self) -> tuple[int, int]:
        ok = sum(1 for a in self.accounts if str(a["status"]).startswith("成功"))
        fail = sum(1 for a in self.accounts if a["status"] == "失败")
        return ok, fail

    def snapshot(self) -> dict[str, Any]:
        with self._lock:
            ok, fail = self.counts()
            return {
                "id": self.id,
                "batch_no": self.batch_no,
                "batch_label": self.batch_label,
                "status": self.status,
                "created_at": self.created_at,
                "count": self.count,
                "concurrency": self.concurrency,
                "dry_run": bool(self.params.get("dry_run")),
                "token_mode": self.params.get("token_mode"),
                "ok_count": ok,
                "fail_count": fail,
                "params": _mask_params(self.params),
                "accounts": [dict(a) for a in self.accounts],
                "logs": list(self.logs[-500:]),
                "batch_summary": self.batch_summary,
            }

    def summary(self) -> dict[str, Any]:
        ok, fail = self.counts()
        return {
            "id": self.id,
            "batch_no": self.batch_no,
            "batch_label": self.batch_label,
            "status": self.status,
            "created_at": self.created_at,
            "count": self.count,
            "concurrency": self.concurrency,
            "dry_run": bool(self.params.get("dry_run")),
            "token_mode": self.params.get("token_mode"),
            "ok_count": ok,
            "fail_count": fail,
        }


class _JobLogHandler(logging.Handler):
    def emit(self, record: logging.LogRecord) -> None:  # noqa: D401
        if not (record.name == "outlook_api_reg" or record.name.startswith("outlook_api_reg.")):
            return
        job_id = _thread_job.get(threading.get_ident()) or _active_batch_job
        if not job_id:
            return
        job = _jobs.get(job_id)
        if not job:
            return
        try:
            msg = record.getMessage()
        except Exception:  # noqa: BLE001
            msg = str(record.msg)
        job.push_log(record.levelname, msg)


def _install_log_handler() -> None:
    handler = _JobLogHandler()
    handler.setLevel(logging.INFO)
    root = logging.getLogger()
    if not any(isinstance(h, _JobLogHandler) for h in root.handlers):
        root.addHandler(handler)
    pkg_logger = logging.getLogger("outlook_api_reg")
    pkg_logger.setLevel(logging.INFO)


_install_log_handler()


def _mask(secret: str) -> str:
    if not secret:
        return ""
    if len(secret) <= 8:
        return "***"
    return secret[:4] + "***" + secret[-4:]


def _mask_params(params: dict[str, Any]) -> dict[str, Any]:
    out = dict(params)
    if out.get("captcha_key"):
        out["captcha_key"] = _mask(out["captcha_key"])
    return out


def _load_jobs_store() -> list[dict[str, Any]]:
    return account_store.list_jobs()


def _job_record(job: "Job") -> dict[str, Any]:
    ok, fail = job.counts()
    emails = [
        a.get("email")
        for a in job.accounts
        if a.get("email") and str(a.get("status") or "").startswith("成功")
    ]
    return {
        "id": job.id,
        "batch_no": job.batch_no,
        "batch_label": job.batch_label,
        "created_at": job.created_at,
        "status": job.status,
        "count": job.count,
        "concurrency": job.concurrency,
        "token_mode": job.params.get("token_mode"),
        "dry_run": bool(job.params.get("dry_run")),
        "ok_count": ok,
        "fail_count": fail,
        "emails": emails,
        "params": _mask_params(job.params),
    }


def _persist_job(job: "Job") -> None:
    if job.params.get("dry_run"):
        return
    account_store.save_job(_job_record(job))


def _batch_index() -> dict[str, dict[str, Any]]:
    out = account_store.batch_index()
    with _jobs_lock:
        jobs = list(_jobs.values())
    for j in jobs:
        info = {"batch_id": j.id, "batch_no": j.batch_no, "batch_label": j.batch_label}
        for a in j.accounts:
            em = a.get("email") or ""
            if em and str(a.get("status") or "").startswith("成功"):
                out[em] = info
    return out


# ---------------------------------------------------------------------------
# 注册执行（并发 + register_batch 特性探测）
# ---------------------------------------------------------------------------


def _run_dry(job: Job) -> None:
    steps = [
        "代理预检通过: (模拟) 出口=203.0.113.7",
        "选用邮箱: 模拟随机前缀@outlook.com",
        "risk/initialize → humanSensorUrl 预加载",
        "captcha.run silent → press（模拟通过）",
        "risk/verify #2 challengeSolution 提交",
        "CreateAccount 成功",
        "oauth20_authorize.srf slt 登录（模拟）",
    ]
    mode = job.params.get("token_mode") or DEFAULT_TOKEN_MODE
    job.push_log("INFO", f"产出格式: {mode}（干跑，仅演示）｜并发度 {job.concurrency}")

    def one(i: int) -> None:
        job.update_account(i, status="进行中")
        job.push_log("INFO", f"[#{i+1}] 干跑开始（不消耗真实资源）")
        for s in steps:
            job.push_log("INFO", f"[#{i+1}] {s}")
            time.sleep(0.08)
        email = f"dryrun{i+1}_{uuid.uuid4().hex[:6]}@outlook.com"
        pwd = f"DryRunPwd{i+1}!"
        if mode in ("graph_recovery", "login_exe", "recovery"):
            rec = f"rec{i+1}_{uuid.uuid4().hex[:6]}@your-cf-domain.com" if mode == "graph_recovery" else f"rec{i+1}_{uuid.uuid4().hex[:6]}@your-recovery-host.com"
            rec_pwd = "cf_domain" if mode == "graph_recovery" else "DryRunRecPwd"
            combo = f"{email}----{pwd}----{MAIL_CLIENT_ID}--------{rec}----{rec_pwd}"
            label = "Graph 六段式" if mode == "graph_recovery" else "login.exe 六段式(IMAP)"
            job.push_log("INFO", f"[#{i+1}] 干跑产出 {label}（恢复邮箱，非 dual）")
        else:
            combo = f"{email}----{pwd}----{MAIL_CLIENT_ID}----"
        job.update_account(
            i, status="成功(干跑)", email=email, password=pwd,
            client_id=MAIL_CLIENT_ID, refresh_token="", combo=combo, error="",
        )
        job.push_log("INFO", f"[#{i+1}] 干跑完成（未写盘、未消耗额度）")

    conc = max(1, min(job.concurrency, job.count))
    with ThreadPoolExecutor(max_workers=conc) as ex:
        for f in as_completed([ex.submit(one, i) for i in range(job.count)]):
            f.result()


def _do_register_one(job: Job, i: int, p: dict[str, Any]) -> None:
    """单个注册（线程池 worker 内执行），日志归属到本任务。"""
    _thread_job[threading.get_ident()] = job.id
    job.update_account(i, status="进行中")
    ph = _proxy_line_for_index(p, i)
    if ph:
        job.push_log("INFO", f"[#{i + 1}] 开始注册（代理 {ph}）")
    else:
        job.push_log("INFO", f"[#{i + 1}] 开始注册")
    plan = p.get("proxy_plan") or []
    one_proxy = plan[i] if i < len(plan) else (p.get("proxy") or None)
    assignments = p.get("proxy_assignments") or []
    try:
        try:
            result: RegisterResult = register_one(
                email_prefix=(p.get("prefix") or None),
                email_domain=p.get("domain") or "@outlook.com",
                country=p.get("country") or "US",
                proxy=one_proxy or None,
                px_mode=p.get("px_mode") or "solver",
                skip_post_login=bool(p.get("skip_login")),
                fetch_mail_token=not bool(p.get("no_mail_token")),
            )
        except Exception as exc:  # noqa: BLE001
            job.update_account(i, status="失败", error=str(exc))
            job.push_log("ERROR", f"[#{i+1}] 异常: {exc}")
            _after_register_proxy(
                "", assignments, i, success=False,
                reg_country=p.get("country") or "US", error=str(exc),
            )
            return
    finally:
        _thread_job.pop(threading.get_ident(), None)

    if result.success:
        saved = ""
        try:
            with _save_lock:
                saved = save_account(
                    result, str(ACCOUNTS_DIR),
                    batch_id=job.id, batch_no=job.batch_no, batch_label=job.batch_label,
                )
        except Exception as exc:  # noqa: BLE001
            job.push_log("ERROR", f"[#{i+1}] 保存失败: {exc}")
        mode = p.get("token_mode") or DEFAULT_TOKEN_MODE
        combo = _job_combo(result, mode)
        job.update_account(
            i, status="成功", email=result.email, password=result.password,
            client_id=result.client_id, refresh_token=result.refresh_token or "",
            recovery_email=result.recovery_email or "",
            recovery_password=result.recovery_password or "",
            combo=combo,
            combo_dual=result.to_combo(dual=True) if result.login_refresh_token else "",
            combo_recovery=result.to_combo(recovery=True) if result.recovery_email else "",
            error="", saved_path=saved,
        )
        tip = "已取 refresh_token" if result.refresh_token else "无 refresh_token"
        if result.recovery_email:
            tip += "，已绑恢复邮箱"
        job.push_log("INFO", f"[#{i+1}] 注册成功 {result.email}（{tip}）")
        _after_register_proxy(
            result.email or "", assignments, i, success=True,
            reg_country=p.get("country") or "US",
        )
    else:
        job.update_account(i, status="失败", error=result.error)
        job.push_log("ERROR", f"[#{i+1}] 注册失败: {result.error}")
        _after_register_proxy(
            "", assignments, i, success=False,
            reg_country=p.get("country") or "US", error=result.error or "",
        )


def _proxy_line_for_index(p: dict[str, Any], idx: int) -> str:
    plan = p.get("proxy_plan") or []
    tpl = plan[idx] if idx < len(plan) else (p.get("proxy") or "")
    if not tpl:
        return ""
    try:
        return proxy_pool.mask_template(str(tpl))
    except Exception:  # noqa: BLE001
        return str(tpl)[:72]


def _run_batch_iter(job: Job, p: dict[str, Any]) -> bool:
    """用引擎 register_batch_iter 驱动 SSE。

    引擎内部已 save_account，网页不再重复保存。返回 True 表示已由本函数处理（成功
    或已消费部分事件不宜重跑），False 表示未开跑可安全回退线程池。
    """
    if not BATCH_READY or register_batch_iter is None:
        return False
    global _active_batch_job
    assignments = p.get("proxy_assignments") or []
    proxy_templates = [
        (a.get("template") or "").strip() or None for a in assignments
    ]
    if len(proxy_templates) < job.count:
        proxy_templates += [None] * (job.count - len(proxy_templates))
    kwargs = dict(
        concurrency=job.concurrency,
        email_prefix=(p.get("prefix") or None),
        email_domain=p.get("domain") or "@outlook.com",
        country=p.get("country") or "US",
        proxy=(p.get("proxy") or None),
        proxy_plan=p.get("proxy_plan"),
        proxy_templates=proxy_templates,
        px_mode=p.get("px_mode") or "solver",
        skip_post_login=bool(p.get("skip_login")),
        fetch_mail_token=not bool(p.get("no_mail_token")),
        output_dir=str(ACCOUNTS_DIR),
        batch_id=job.id,
        batch_no=job.batch_no,
        batch_label=job.batch_label,
        jitter_min=p.get("jitter_min"),
        jitter_max=p.get("jitter_max"),
    )
    _active_batch_job = job.id  # 引擎线程池内 register_one 的日志归属到本任务
    consumed = False
    try:
        for ev in register_batch_iter(job.count, **kwargs):
            etype = ev.get("type")
            if etype == "start":
                jitter = ev.get("jitter") or []
                jtxt = (
                    f"{jitter[0]}–{jitter[1]}秒"
                    if len(jitter) >= 2 and float(jitter[1] or 0) > 0
                    else "无"
                )
                job.push_log(
                    "INFO",
                    f"引擎批量启动：共 {ev.get('total')} 个，并发 {ev.get('concurrency')}，"
                    f"相邻启动错峰 {jtxt}",
                )
                if ev.get("proxy_unique"):
                    job.push_log("INFO", "防封·一号一 IP：每号独立 {sid} 会话（出口 IP 不同）")
                elif not ev.get("proxy_has_sid") and (p.get("proxy") or p.get("proxy_plan")):
                    job.push_log(
                        "WARNING",
                        "代理未含 {sid}：全批可能共用同一出口 IP，建议改用带 {sid} 的模板",
                    )
            elif etype == "account_start":
                consumed = True
                idx = int(ev.get("index", 0))
                if idx >= job.count:
                    continue
                job.update_account(idx, status="进行中")
                ph = _proxy_line_for_index(p, idx)
                if ph:
                    job.push_log("INFO", f"[#{idx + 1}] 开始注册（代理 {ph}）")
                else:
                    job.push_log("INFO", f"[#{idx + 1}] 开始注册")
            elif etype == "result":
                consumed = True
                idx = int(ev.get("index", 0))
                if idx >= job.count:
                    continue
                i = idx
                combo = ev.get("combo") or ""
                rec_combo = ev.get("combo_recovery") or ""
                if p.get("token_mode") in ("login_exe", "recovery", "graph_recovery") and rec_combo:
                    combo = rec_combo
                if combo:
                    _e, pwd, cid, rt = _split_combo(combo)
                else:
                    pwd = cid = rt = ""
                if ev.get("success"):
                    job.update_account(
                        i, status="成功", email=ev.get("email") or "",
                        password=pwd, client_id=cid, refresh_token=rt,
                        recovery_email=ev.get("recovery_email") or "",
                        recovery_password=ev.get("recovery_password") or "",
                        combo=combo, combo_dual=ev.get("combo_dual") or "",
                        combo_recovery=rec_combo,
                        login_token=bool(ev.get("login_token_present")),
                        error="", saved_path="(引擎已保存)",
                    )
                    _after_register_proxy(
                        ev.get("email") or "", assignments, idx, success=True,
                        reg_country=p.get("country") or "US",
                    )
                    dual_tip = "，含双令牌" if ev.get("login_token_present") else ""
                    job.push_log(
                        "INFO",
                        f"[#{i+1}] 成功 {ev.get('email')}（{ev.get('elapsed')}s{dual_tip}）",
                    )
                else:
                    job.update_account(
                        i, status="失败", email=ev.get("email") or "",
                        error=ev.get("error") or "",
                    )
                    _after_register_proxy(
                        ev.get("email") or "", assignments, idx, success=False,
                        reg_country=p.get("country") or "US",
                        error=ev.get("error") or "",
                    )
                    job.push_log("ERROR", f"[#{i+1}] 失败：{ev.get('error')}")
            elif etype == "done":
                job.batch_summary = {
                    "total": ev.get("total"), "ok": ev.get("ok"),
                    "failed": ev.get("failed"), "elapsed": ev.get("elapsed"),
                    "avg_per_account": ev.get("avg_per_account"),
                    "avg_stage_timings": ev.get("avg_stage_timings") or {},
                }
                job.emit({"type": "summary", "summary": job.batch_summary})
                stages = job.batch_summary["avg_stage_timings"]
                top = sorted(stages.items(), key=lambda kv: kv[1], reverse=True)[:3]
                top_txt = "，".join(f"{k}={v}s" for k, v in top) or "无"
                job.push_log(
                    "INFO",
                    f"本批完成：成功 {ev.get('ok')}/{ev.get('total')}，本批耗时 "
                    f"{ev.get('elapsed')}s，单号均耗 {ev.get('avg_per_account')}s，"
                    f"阶段大头：{top_txt}",
                )
        return True
    except Exception as exc:  # noqa: BLE001
        job.push_log("WARNING", f"引擎 register_batch_iter 执行异常: {exc}")
        return consumed  # 已消费部分事件则不重跑，避免重复真实注册
    finally:
        _active_batch_job = None


def _run_real(job: Job) -> None:
    p = job.params
    # DB 为主：任务带了 key 就用并回落库；没带则从库里回读注入环境变量。
    if p.get("captcha_key"):
        os.environ["CAPTCHA_RUN_API_KEY"] = p["captcha_key"]
        app_db.set_setting("CAPTCHA_RUN_API_KEY", p["captcha_key"])
    elif not os.environ.get("CAPTCHA_RUN_API_KEY"):
        db_key = app_db.get_setting("CAPTCHA_RUN_API_KEY")
        if db_key:
            os.environ["CAPTCHA_RUN_API_KEY"] = db_key

    job.push_log("INFO", "正在规划代理…")
    try:
        plan, pmeta = _build_register_proxy_plan(p, job.count)
    except ValueError as exc:
        job.push_log("ERROR", str(exc))
        raise
    p["proxy_plan"] = plan
    pmeta_assignments = pmeta.get("assignments") or []
    if not pmeta_assignments:
        manual_tpl = (p.get("proxy") or "").strip()
        pmeta_assignments = [
            {
                "index": i,
                "template": manual_tpl,
                "resolved": plan[i] if i < len(plan) else "",
            }
            for i in range(job.count)
        ]
    p["proxy_assignments"] = pmeta_assignments
    if p.get("use_proxy_pool"):
        job.push_log(
            "INFO",
            f"代理池：已规划 {len(plan)} 条（策略 {pmeta.get('strategy') or '—'}）",
        )
    retries = max(1, int((os.environ.get("REG_PROXY_RETRIES") or "3").strip() or "3"))
    job.push_log("INFO", f"PX/代理失败自动重试：最多 {retries} 次（REG_PROXY_RETRIES）")

    requested_mode = p.get("token_mode") or DEFAULT_TOKEN_MODE
    effective = _apply_token_mode(requested_mode)
    if effective != requested_mode:
        job.push_log("WARNING", f"产出格式 {requested_mode} 不可用，已降级为 {effective}")
    fmt_label = {"graph": "Graph 四段", "graph_recovery": "Graph 六段", "dual": "双令牌六段"}.get(
        effective, effective
    )
    job.push_log("INFO", f"产出格式: {fmt_label}｜并发度 {job.concurrency}")

    # 优先引擎生成器；不可用则线程池兜底
    if _run_batch_iter(job, p):
        return
    job.push_log("INFO", "register_batch_iter 不可用，使用线程池兜底并发 register_one")
    conc = max(1, min(job.concurrency, job.count))
    with ThreadPoolExecutor(max_workers=conc) as ex:
        futs = [ex.submit(_do_register_one, job, i, p) for i in range(job.count)]
        for f in as_completed(futs):
            f.result()


def _job_worker(job_id: str) -> None:
    job = _jobs[job_id]
    _thread_job[threading.get_ident()] = job_id
    job.push_log(
        "INFO",
        f"批次 {job.batch_label} 开始执行（{job.count} 个，并发 {job.concurrency}）",
    )
    try:
        if job.params.get("dry_run"):
            _run_dry(job)
        else:
            _run_real(job)
        job.status = "done"
    except Exception as exc:  # noqa: BLE001
        job.status = "error"
        job.push_log("ERROR", f"任务异常终止: {exc}")
    finally:
        _thread_job.pop(threading.get_ident(), None)
        try:
            _persist_job(job)
        except Exception:  # noqa: BLE001
            pass
        job.emit({"type": "done", "status": job.status})
        job.push_log("INFO", "任务结束。")


# ---------------------------------------------------------------------------
# 请求模型
# ---------------------------------------------------------------------------


class RegisterRequest(BaseModel):
    count: int = 1
    concurrency: int = 2
    prefix: Optional[str] = None
    domain: str = "@outlook.com"
    country: str = "US"
    proxy: Optional[str] = None
    px_mode: str = "solver"
    skip_login: bool = False
    no_mail_token: bool = False
    captcha_key: Optional[str] = None
    token_mode: str = DEFAULT_TOKEN_MODE
    dry_run: bool = False
    # 防封·启动错峰（秒）：留空则用引擎默认（3~8）/ 环境变量 OUTLOOK_REG_JITTER_MIN/MAX
    jitter_min: Optional[float] = None
    jitter_max: Optional[float] = None
    batch_label: Optional[str] = None  # 留空则按 日期-国家-域名-数量-格式 自动生成
    use_proxy_pool: bool = False


class ProxyPoolAddRequest(BaseModel):
    templates: list[str] = []
    text: Optional[str] = None  # 多行批量导入
    label: Optional[str] = None
    provider: Optional[str] = None  # 代理商/供应商分组
    country: Optional[str] = None  # 代理出口国家 ISO2，留空则从模板推断


class ProxyPoolUpdateRequest(BaseModel):
    label: Optional[str] = None
    template: Optional[str] = None
    provider: Optional[str] = None
    country: Optional[str] = None
    enabled: Optional[bool] = None


class ProxyPoolDeleteRequest(BaseModel):
    ids: list[str]


class ProxyPoolCheckRequest(BaseModel):
    ids: Optional[list[str]] = None
    timeout: int = 15


class ProxyPoolSettingsRequest(BaseModel):
    strategy: Optional[str] = None
    require_healthy: Optional[bool] = None
    sticky_per_account: Optional[bool] = None


class ProxyPoolBindRequest(BaseModel):
    email: str
    proxy_id: str


class ProxyPoolUnbindRequest(BaseModel):
    emails: list[str]


class ProxyPoolEnsureRequest(BaseModel):
    templates: list[str] = []
    text: Optional[str] = None
    provider: Optional[str] = "web"
    country: Optional[str] = None


class VerifyComboRequest(BaseModel):
    combo: Optional[str] = None
    email: Optional[str] = None
    refresh_token: Optional[str] = None
    proxy: Optional[str] = None
    test_imap: bool = True


class VerifyBatchRequest(BaseModel):
    emails: Optional[list[str]] = None  # 指定账号邮箱；空=全部有 token 的
    combos: Optional[list[str]] = None  # 或直接传 combo 列表
    proxy: Optional[str] = None
    test_imap: bool = False
    concurrency: int = 4


class ImportRequest(BaseModel):
    text: str = ""


class ExportRequest(BaseModel):
    emails: Optional[list[str]] = None
    format: str = "graph"


class DeleteRequest(BaseModel):
    emails: list[str]


class MetaRequest(BaseModel):
    email: str
    note: Optional[str] = None
    tags: Optional[list[str]] = None


class KeepaliveRequest(BaseModel):
    emails: Optional[list[str]] = None  # 选中账号邮箱；空=全部有 token 的
    proxy: Optional[str] = None
    concurrency: int = 5


class RescueRequest(BaseModel):
    emails: Optional[list[str]] = None
    proxy: Optional[str] = None
    concurrency: int = 1
    use_proxy_pool: bool = False


class ReplenishRequest(BaseModel):
    emails: Optional[list[str]] = None
    proxy: Optional[str] = None
    verify: bool = True


# ---------------------------------------------------------------------------
# 静态页 / 配置
# ---------------------------------------------------------------------------


@app.on_event("startup")
def _startup_log() -> None:
    try:
        app_db.ensure_initialized(ACCOUNTS_DIR)
        st = app_db.db_status()
        logger.info(
            "SQLite 已就绪: %s（账号 %s · 代理 %s）",
            st.get("path"),
            st.get("accounts", 0),
            st.get("proxies", 0),
        )
    except Exception as exc:  # noqa: BLE001
        logger.exception("SQLite 初始化失败: %s", exc)
    if coolhs_mail.coolhs_backend_active() and coolhs_mail.coolhs_configured():
        logging.getLogger(__name__).info(
            "proofs 收码后端: coolhs-mail %s (%s)",
            coolhs_mail.load_config().domain,
            coolhs_mail.load_config().base_url,
        )
    elif cf_domain_mail.cf_domain_backend_active() and cf_domain_mail.cf_configured():
        logging.getLogger(__name__).info(
            "proofs 收码后端: CF 域名 %s", cf_domain_mail.load_config().domain
        )
    elif not ext_recovery_pool.external_pool_enabled():
        logging.getLogger(__name__).warning(
            "恢复邮箱未配置：请设置 OUTLOOK_RECOVERY_BACKEND=coolhs_mail"
            "（COOLHS_MAIL_*）或 cf_domain，或 IMAP 池 OUTLOOK_EXTERNAL_RECOVERY_POOL_FILE"
        )


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse(content=(STATIC_DIR / "index.html").read_text(encoding="utf-8"))


@app.get("/api/config")
def get_config() -> JSONResponse:
    # 代理不回显；captcha.run key 以数据库为主，回显「是否已存 + 掩码」，不回明文。
    proxy = ""
    captcha_key = (os.environ.get("CAPTCHA_RUN_API_KEY") or app_db.get_setting("CAPTCHA_RUN_API_KEY") or "").strip()
    cf_active = cf_domain_mail.cf_domain_backend_active() and cf_domain_mail.cf_configured()
    coolhs_active = coolhs_mail.coolhs_backend_active() and coolhs_mail.coolhs_configured()
    recovery_pool_configured = ext_recovery_pool.external_pool_enabled()
    recovery_configured = cf_active or coolhs_active or recovery_pool_configured
    recovery_backend = cf_domain_mail.recovery_backend()
    env_mode = DEFAULT_TOKEN_MODE
    if env_mode in ("login_exe", "recovery"):
        default_product = "graph_recovery"
    elif cf_active or coolhs_active or env_mode == "graph":
        default_product = "graph_recovery"
    else:
        default_product = "graph"
    return JSONResponse(
        {
            "proxy": proxy,
            "captcha_key_masked": _mask(captcha_key),
            "captcha_key_set": bool(captcha_key),
            "px_modes": ["solver"],
            "domains": [
                "@outlook.com",
                "@hotmail.com",
                "@outlook.com.au",
                "@outlook.de",
                "@outlook.jp",
            ],
            "default_country": "US",
            "mail_client_id": MAIL_CLIENT_ID,
            "token_modes": TOKEN_MODES,
            "product_modes": PRODUCT_MODES,
            "default_token_mode": default_product,
            "mail_token_mode_env": env_mode,
            "dual_ready": DUAL_READY,
            "batch_ready": BATCH_READY,
            "keepalive_ready": KEEPALIVE_READY,
            "rescue_ready": RESCUE_READY,
            "export_formats": EXPORT_FORMATS,
            "recovery_pool_configured": recovery_pool_configured,
            "cf_domain_configured": cf_active,
            "coolhs_mail_configured": coolhs_active,
            "recovery_backend": recovery_backend,
            "recovery_configured": recovery_configured,
            "proxy_pool": proxy_pool.pool_stats(),
            "proxy_pool_file": str(proxy_pool.pool_file()),
            "proxy_pool_backend": proxy_pool.storage_backend(),
            "database": app_db.db_status(),
        }
    )


# ---------------------------------------------------------------------------
# 应用设置（DB 为主：captcha.run / EzCaptcha / CapSolver key 存 app_meta）
# ---------------------------------------------------------------------------

_SETTINGS_KEYS = ("CAPTCHA_RUN_API_KEY", "EZCAPTCHA_API_KEY", "CAPSOLVER_API_KEY")


class SettingsRequest(BaseModel):
    captcha_run_api_key: Optional[str] = None
    ezcaptcha_api_key: Optional[str] = None
    capsolver_api_key: Optional[str] = None


def _setting_status(key: str) -> dict[str, Any]:
    val = (os.environ.get(key) or app_db.get_setting(key) or "").strip()
    return {"set": bool(val), "masked": _mask(val), "source": "env" if os.environ.get(key) else ("db" if val else "")}


@app.get("/api/settings")
def get_settings() -> JSONResponse:
    return JSONResponse({k.lower(): _setting_status(k) for k in _SETTINGS_KEYS})


@app.post("/api/settings")
def save_settings(req: SettingsRequest) -> JSONResponse:
    """把打码服务 key 存进数据库（DB 为主）。传空串=清空该项；不传=保持不变。"""
    mapping = {
        "CAPTCHA_RUN_API_KEY": req.captcha_run_api_key,
        "EZCAPTCHA_API_KEY": req.ezcaptcha_api_key,
        "CAPSOLVER_API_KEY": req.capsolver_api_key,
    }
    changed = []
    for key, value in mapping.items():
        if value is None:
            continue
        app_db.set_setting(key, value.strip())
        # 同步进程环境变量，使本进程内即时生效（引擎/CLI 复用同进程时无需重启）。
        if value.strip():
            os.environ[key] = value.strip()
        else:
            os.environ.pop(key, None)
        changed.append(key.lower())
    return JSONResponse({"ok": True, "changed": changed, "settings": {k.lower(): _setting_status(k) for k in _SETTINGS_KEYS}})


# ---------------------------------------------------------------------------
# 注册 / 任务
# ---------------------------------------------------------------------------


@app.post("/api/register")
def start_register(req: RegisterRequest) -> JSONResponse:
    if req.count < 1:
        raise HTTPException(status_code=400, detail="数量至少为 1。")
    if not req.dry_run and req.count > 20:
        raise HTTPException(status_code=400, detail="真实注册单次上限 20，请分批。")
    if not req.dry_run:
        # DB 为主：页面填了就落库；没填则回读库里已存的 key。
        provided_key = (req.captcha_key or "").strip()
        if provided_key:
            app_db.set_setting("CAPTCHA_RUN_API_KEY", provided_key)
        captcha_key = (
            provided_key
            or os.environ.get("CAPTCHA_RUN_API_KEY")
            or app_db.get_setting("CAPTCHA_RUN_API_KEY")
            or ""
        ).strip()
        proxy = (req.proxy or "").strip()
        if proxy:
            proxy_pool.ensure_templates([proxy], provider="web")
        if req.use_proxy_pool:
            stats = proxy_pool.pool_stats()
            if stats.get("enabled", 0) < 1 and not proxy:
                raise HTTPException(
                    status_code=400,
                    detail="代理池为空。请在「代理池」页添加，或在注册页填写代理（会自动写入数据库）。",
                )
        elif not proxy:
            raise HTTPException(status_code=400, detail="请填写代理或启用「使用代理池」。")
        if not captcha_key:
            raise HTTPException(status_code=400, detail="请填写 captcha.run Key（Web 页对应输入框，会存入数据库，下次免填）。")
    concurrency = max(1, min(int(req.concurrency or 1), req.count))

    with _jobs_lock:
        running_real = [
            j for j in _jobs.values() if j.status == "running" and not j.params.get("dry_run")
        ]
        if running_real and not req.dry_run:
            raise HTTPException(status_code=409, detail="已有注册任务进行中，请等待其完成。")
        job_id = uuid.uuid4().hex[:12]
        params = req.model_dump()
        params["concurrency"] = concurrency
        stored_nos = [int(r.get("batch_no") or 0) for r in _load_jobs_store()]
        live_nos = [int(getattr(j, "batch_no", 0) or 0) for j in _jobs.values()]
        params["batch_no"] = max(stored_nos + live_nos + [0]) + 1
        params["batch_label"] = _make_batch_label(params, params["batch_no"], jobs=_jobs)
        job = Job(job_id, params)
        _jobs[job_id] = job
        try:
            _persist_job(job)
        except Exception:  # noqa: BLE001
            pass

    threading.Thread(target=_job_worker, args=(job_id,), daemon=True).start()
    return JSONResponse(
        {
            "job_id": job_id,
            "count": req.count,
            "concurrency": concurrency,
            "dry_run": req.dry_run,
            "batch_no": params["batch_no"],
            "batch_label": params["batch_label"],
        }
    )


@app.get("/api/jobs")
def list_jobs() -> JSONResponse:
    merged: dict[str, dict[str, Any]] = {}
    for rec in _load_jobs_store():
        if rec.get("id"):
            merged[rec["id"]] = rec
    with _jobs_lock:
        for j in _jobs.values():
            merged[j.id] = j.summary()
    jobs = sorted(merged.values(), key=lambda x: x.get("created_at") or "", reverse=True)
    return JSONResponse({"jobs": jobs})


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str) -> JSONResponse:
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="任务不存在")
    return JSONResponse(job.snapshot())


@app.get("/api/jobs/{job_id}/events")
def job_events(job_id: str) -> StreamingResponse:
    job = _jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="任务不存在")

    def gen():
        yield f"data: {json.dumps({'type': 'snapshot', 'snapshot': job.snapshot()}, ensure_ascii=False)}\n\n"
        while True:
            try:
                ev = job._queue.get(timeout=15)
            except queue.Empty:
                yield ": keep-alive\n\n"
                if job.status != "running":
                    break
                continue
            yield f"data: {json.dumps(ev, ensure_ascii=False)}\n\n"
            if ev.get("type") == "done":
                break

    return StreamingResponse(
        gen(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ---------------------------------------------------------------------------
# 账号读取 / 统计 / 导入导出 / 删除 / 元数据
# ---------------------------------------------------------------------------


def _split_combo(combo: str) -> tuple[str, str, str, str]:
    """email----password----client_id----refresh_token → (email, pwd, cid, rt)。"""
    parts = combo.strip().split("----")
    if len(parts) < 4:
        return "", "", "", ""
    email = parts[0]
    pwd = parts[1]
    rt = next((p for p in parts if p.startswith("M.C")), "") or (parts[3] if len(parts) > 3 else "")
    cid = next((p for p in parts if len(p) == 36 and p.count("-") == 4), "") or parts[2]
    return email, pwd, cid, rt


def _apply_batch(row: dict[str, Any], data: dict[str, Any], index: dict[str, dict[str, Any]]) -> None:
    email = str(row.get("email") or "")
    hit = index.get(email) or {}
    batch_id = data.get("batch_id") or hit.get("batch_id") or ""
    batch_no = data.get("batch_no") if data.get("batch_no") not in (None, "") else hit.get("batch_no")
    batch_label = data.get("batch_label") or hit.get("batch_label") or (
        f"B{batch_no}" if batch_no else ""
    )
    row["batch_id"] = batch_id
    row["batch_no"] = batch_no
    row["batch_label"] = batch_label


def _load_accounts() -> list[dict[str, Any]]:
    """账号池列表：SQLite 为唯一数据源。"""
    app_db.ensure_initialized(ACCOUNTS_DIR)
    rows = account_store.list_accounts()
    batch_map = _batch_index()
    for row in rows:
        m = {"verify": row.get("verify")}
        _apply_batch(row, {
            "batch_id": row.get("batch_id"),
            "batch_no": row.get("batch_no"),
            "batch_label": row.get("batch_label"),
        }, batch_map)
        if not row.get("updated_at"):
            row["updated_at"] = _best_updated_at(row, m)
        if not row.get("last_alive_at"):
            row["last_alive_at"] = _best_last_alive(row, m)
        end_dt = _survival_end_at(row, m)
        row["survival_end_at"] = _dt_iso(end_dt) if end_dt else ""
        row["alive_seconds"] = _compute_alive_seconds(str(row.get("created_at") or ""), end_dt)
    return rows


def _is_graph_readable(row: dict[str, Any]) -> bool:
    v = row.get("verify") or {}
    return bool(v.get("ok") or v.get("graph"))


def _compute_stats(rows: list[dict[str, Any]]) -> dict[str, Any]:
    today = datetime.now().strftime("%Y-%m-%d")
    total = len(rows)
    with_token = sum(1 for r in rows if r.get("has_token"))
    usable = dead = untested = 0
    today_new = 0
    for r in rows:
        v = r.get("verify")
        if v is None:
            untested += 1
        elif _is_graph_readable(r):
            usable += 1
        elif r.get("batch_label") or r.get("batch_no"):
            dead += 1
        # 无批次旧号的失败测活：概览不计入失活/未测，避免数字膨胀
        if r.get("created_at", "").startswith(today):
            today_new += 1
    return {
        "total": total,
        "with_token": with_token,
        "usable": usable,
        "dead": dead,
        "untested": untested,
        "today_new": today_new,
        "recovery_pool_configured": ext_recovery_pool.external_pool_enabled(),
    }


@app.get("/api/accounts")
def list_accounts() -> JSONResponse:
    rows = _load_accounts()
    return JSONResponse({"count": len(rows), "accounts": rows, "stats": _compute_stats(rows)})


def _format_combo(row: dict[str, Any], fmt: str) -> str:
    if fmt == "dual":
        return row.get("combo_dual") or row.get("combo", "")
    if fmt == "recovery":
        if row.get("combo_recovery"):
            return row["combo_recovery"]
        rec_email = row.get("recovery_email", "")
        rec_pwd = row.get("recovery_password", "")
        if rec_email and rec_pwd and row.get("combo"):
            parts = row["combo"].split("----")
            if len(parts) >= 4:
                return "----".join(parts[:4] + [rec_email, rec_pwd])
        return ""
    return row.get("combo", "")


@app.get("/api/accounts/export")
def export_accounts_all() -> PlainTextResponse:
    rows = _load_accounts()
    body = "\n".join(r["combo"] for r in rows if r.get("combo"))
    body += "\n" if body else ""
    return PlainTextResponse(body, headers={"Content-Disposition": "attachment; filename=accounts_combo.txt"})


@app.post("/api/accounts/export")
def export_accounts(req: ExportRequest) -> PlainTextResponse:
    fmt = req.format if req.format in EXPORT_FORMATS else "graph"
    rows = _load_accounts()
    if req.emails:
        want = set(req.emails)
        rows = [r for r in rows if r["email"] in want]
    lines: list[str] = []
    six_count = 0  # 实际输出为 6 段的行数（dual/recovery 且有对应字段）
    for r in rows:
        c = _format_combo(r, fmt)
        if not c:
            continue
        lines.append(c)
        if fmt == "dual" and r.get("combo_dual"):
            six_count += 1
        if fmt == "recovery" and len(c.split("----")) >= 6:
            six_count += 1
    body = "\n".join(lines) + ("\n" if lines else "")
    total = len(lines)
    degraded_full = fmt in ("dual", "recovery") and total > 0 and six_count == 0
    return PlainTextResponse(
        body,
        headers={
            "Content-Disposition": f"attachment; filename=accounts_{fmt}.txt",
            "X-Export-Degraded": "1" if degraded_full else "0",
            "X-Export-Six": str(six_count),
            "X-Export-Total": str(total),
            "Access-Control-Expose-Headers": "X-Export-Degraded,X-Export-Six,X-Export-Total",
        },
    )


@app.post("/api/accounts/import")
def import_accounts(req: ImportRequest) -> JSONResponse:
    """自动识别 4 段/6 段：均写入 accounts.txt（graph 四段）；6 段额外写 accounts_dual.txt
    并把登录令牌存进 meta，便于之后 6 段导出。按邮箱去重、校验字段。"""
    existing = {r["email"] for r in _load_accounts()}
    imported = duplicate = invalid = six_seg = 0
    seen: set[str] = set()
    graph_lines: list[str] = []
    dual_lines: list[str] = []
    dual_meta: dict[str, dict[str, str]] = {}
    for raw in (req.text or "").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("----")
        if len(parts) < 4:
            invalid += 1
            continue
        email = parts[0].strip()
        if not email or "@" not in email:
            invalid += 1
            continue
        if email in existing or email in seen:
            duplicate += 1
            continue
        seen.add(email)
        graph_lines.append("----".join(parts[:4]))
        imported += 1
        # 6 段：email----pwd----graph_cid----graph_rt----login_cid----login_rt
        if len(parts) >= 6 and parts[4].strip() and parts[5].strip():
            six = "----".join(parts[:6])
            dual_lines.append(six)
            dual_meta[email] = {
                "combo_dual": six,
                "login_client_id": parts[4].strip(),
                "login_refresh_token": parts[5].strip(),
            }
            six_seg += 1
    if graph_lines or dual_lines:
        now = datetime.now().isoformat()
        with _save_lock:
            conn = app_db.connect()
            try:
                for ln in graph_lines:
                    parts = ln.split("----")
                    if len(parts) < 4:
                        continue
                    email = parts[0].strip()
                    dual = dual_meta.get(email, {})
                    account_store.upsert_account_dict(conn, {
                        "email": email,
                        "password": parts[1],
                        "client_id": parts[2],
                        "refresh_token": parts[3],
                        "combo": ln,
                        "combo_dual": dual.get("combo_dual", ""),
                        "login_client_id": dual.get("login_client_id", ""),
                        "login_refresh_token": dual.get("login_refresh_token", ""),
                        "success": True,
                        "created_at": now,
                        "updated_at": now,
                        "batch_id": "import",
                        "batch_label": "导入",
                        "legacy_source": "import",
                    })
                conn.commit()
            finally:
                conn.close()
    for email, patch in dual_meta.items():
        _update_meta(email, patch)
    return JSONResponse(
        {"ok": True, "imported": imported, "duplicate": duplicate,
         "invalid": invalid, "six_seg": six_seg}
    )


@app.post("/api/accounts/delete")
def delete_accounts(req: DeleteRequest) -> JSONResponse:
    targets = [e for e in req.emails if e]
    if not targets:
        raise HTTPException(status_code=400, detail="未指定要删除的账号。")
    removed = account_store.delete_accounts(targets)
    return JSONResponse({"ok": True, "removed": removed})


@app.post("/api/accounts/meta")
def set_meta(req: MetaRequest) -> JSONResponse:
    patch: dict[str, Any] = {}
    if req.note is not None:
        patch["note"] = req.note
    if req.tags is not None:
        patch["tags"] = req.tags
    if not patch:
        raise HTTPException(status_code=400, detail="无可更新字段。")
    _update_meta(req.email, patch)
    return JSONResponse({"ok": True, "email": req.email, **patch})


# ---------------------------------------------------------------------------
# 可用性校验（单条 + 批量）
# ---------------------------------------------------------------------------


def _verify_one(email: str, refresh_token: str, proxy_url: str, test_imap: bool) -> dict[str, Any]:
    if not refresh_token:
        return {"ok": False, "email": email, "usable": [], "message": "缺少 refresh_token"}

    def _probe(via: str) -> dict[str, Any]:
        return graph_mail.probe_token(email or "unknown", refresh_token, proxy_url=via)

    probe = _probe(proxy_url)
    used_proxy = proxy_url
    if probe.get("transient") and proxy_url:
        logger.warning(
            "测活经代理 %s 网络失败，回退直连重试（refresh 换票无需注册代理）",
            proxy_url.split("@")[-1][:40] if "@" in proxy_url else proxy_url[:40],
        )
        probe = _probe("")
        used_proxy = ""
    if probe.get("transient"):
        detail = probe.get("detail", {})
        return {
            "ok": False,
            "transient": True,
            "unable": True,
            "email": email,
            "usable": [],
            "granted_scope": detail.get("granted_scope", ""),
            "refresh_error": detail.get("refresh", ""),
            "graph": {"status": detail.get("graph"), "ok": False},
            "outlook_rest": {"status": detail.get("outlook_rest"), "ok": False},
            "imap": {"tested": False},
            "summary": "测活暂不可用（网络/SSL），未改账号状态",
            "message": detail.get("refresh", "network"),
            "verify_via": "direct" if not used_proxy else "proxy",
        }
    detail = probe.get("detail", {})
    usable = list(probe.get("usable", []))
    graph_status = detail.get("graph")
    rest_status = detail.get("outlook_rest")
    res: dict[str, Any] = {
        "email": email,
        "granted_scope": detail.get("granted_scope", ""),
        "refresh_error": detail.get("refresh", ""),
        "graph": {"status": graph_status, "ok": graph_status == 200},
        "outlook_rest": {"status": rest_status, "ok": rest_status == 200},
        "imap": {"tested": False},
        "usable": usable,
        "verify_via": "direct" if not used_proxy else "proxy",
    }
    if proxy_url and not used_proxy:
        res["proxy_fallback"] = True
    if test_imap and email:
        im = enable_imap.imap_login_test(email, refresh_token, proxy_url=used_proxy)
        res["imap"] = {
            "tested": True,
            "ok": bool(im.get("ok")),
            "stage": im.get("stage", ""),
            "detail": str(im.get("detail", ""))[:200],
            "message_count": im.get("message_count"),
        }
        if im.get("ok"):
            usable.append("imap")
    res["ok"] = bool(usable)
    if "graph" in usable:
        res["summary"] = "✅ 可用：Graph 令牌可读信（推荐）"
    elif "outlook_rest" in usable:
        res["summary"] = "✅ 可用：Outlook REST 令牌可读信"
    elif res["imap"].get("ok"):
        res["summary"] = "⚠️ 仅 IMAP 可用（老号）"
    else:
        res["summary"] = "❌ 不可用：graph/outlook_rest/imap 均未通过"
    return res


def _cache_verify(res: dict[str, Any]) -> None:
    email = res.get("email")
    if not email:
        return
    account_store.cache_verify(email, res)


@app.post("/api/verify-combo")
def verify_combo(req: VerifyComboRequest) -> JSONResponse:
    email = (req.email or "").strip()
    refresh_token = (req.refresh_token or "").strip()
    if req.combo:
        c_email, _pwd, _cid, c_rt = _split_combo(req.combo)
        email = email or c_email
        refresh_token = refresh_token or c_rt
    if not refresh_token:
        return JSONResponse({"ok": False, "email": email, "usable": [],
                             "message": "缺少 refresh_token（该 combo 第四段为空，无法校验）。"})
    try:
        res = _verify_one(email, refresh_token, _proxy_url(req.proxy), req.test_imap)
        try:
            _cache_verify(res)
        except Exception as exc:  # noqa: BLE001
            logger.exception("缓存测活结果失败: %s", exc)
        return JSONResponse(res)
    except Exception as exc:  # noqa: BLE001
        logger.exception("测活异常 %s: %s", email, exc)
        transient = graph_mail.is_transient_error(exc)
        return JSONResponse({
            "ok": False,
            "transient": transient,
            "unable": transient,
            "email": email,
            "usable": [],
            "summary": (
                "测活暂不可用（网络/SSL），未改账号状态"
                if transient
                else f"测活异常: {exc}"[:160]
            ),
            "message": str(exc)[:160],
        })


@app.post("/api/verify-batch")
def verify_batch(req: VerifyBatchRequest) -> JSONResponse:
    proxy_url = _proxy_url(req.proxy)
    tasks: list[tuple[str, str]] = []  # (email, refresh_token)
    if req.combos:
        for c in req.combos:
            e, _p, _c, rt = _split_combo(c)
            if e:
                tasks.append((e, rt))
    else:
        rows = _load_accounts()
        want = set(req.emails) if req.emails else None
        for r in rows:
            if want is not None and r["email"] not in want:
                continue
            if not r.get("refresh_token"):
                continue
            tasks.append((r["email"], r["refresh_token"]))

    if not tasks:
        return JSONResponse({"ok": True, "results": [], "message": "无可校验账号（缺 refresh_token）。"})

    conc = max(1, min(int(req.concurrency or 4), 8, len(tasks)))
    results: list[dict[str, Any]] = []

    def work(item: tuple[str, str]) -> dict[str, Any]:
        try:
            r = _verify_one(item[0], item[1], proxy_url, req.test_imap)
        except Exception as exc:  # noqa: BLE001
            transient = graph_mail.is_transient_error(exc)
            r = {
                "ok": False,
                "transient": transient,
                "unable": transient,
                "email": item[0],
                "usable": [],
                "summary": (
                    "测活暂不可用（网络/SSL），未改账号状态"
                    if transient
                    else f"测活异常: {exc}"[:160]
                ),
                "message": str(exc)[:160],
            }
        if not r.get("transient"):
            _cache_verify(r)
        return r

    with ThreadPoolExecutor(max_workers=conc) as ex:
        for f in as_completed([ex.submit(work, t) for t in tasks]):
            results.append(f.result())

    ok_n = sum(1 for r in results if r.get("ok"))
    return JSONResponse({"ok": True, "total": len(results), "usable": ok_n, "results": results})


# ---------------------------------------------------------------------------
# IMAP / 保活（占位，如实反馈）
# ---------------------------------------------------------------------------


@app.post("/api/imap-enable")
def imap_enable() -> JSONResponse:
    return JSONResponse(
        {
            "ok": False,
            "implemented": False,
            "required": False,
            "message": "开启 IMAP 为可选项，非必需：默认走 Graph 令牌读信，不依赖 IMAP 协议开关。"
            "主动开启（SetConsumerMailbox）需网页会话 OWA usertoken，纯 API 链路暂未产出；"
            "且新号会返回 412（反滥用），约 10–24h 账号成熟后才可能开成。"
            "确认某号 IMAP 状态请用『测活』勾选 IMAP。",
        }
    )


def _account_json_path_for(email: str) -> Optional[Path]:
    if not ACCOUNTS_DIR.exists():
        return None
    skip = {META_FILE.name, JOBS_FILE.name}
    for fp in ACCOUNTS_DIR.glob("*.json"):
        if fp.name in skip:
            continue
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            continue
        if isinstance(data, dict) and data.get("email") == email:
            return fp
    return None


def _writeback_keepalive(email: str, new_line: str) -> None:
    """把轮换后的新行回写 账号 json / accounts.txt / accounts_dual.txt（在 _save_lock 内调用）。"""
    parts = new_line.split("----")
    if len(parts) < 4:
        return
    graph4 = "----".join(parts[:4])
    new_rt = parts[3]
    is_dual = len(parts) >= 6
    # 账号 json：更新 graph refresh_token / combo，六段时同步 combo_dual / login_*
    fp = _account_json_path_for(email)
    if fp:
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
            data["refresh_token"] = new_rt
            data["combo"] = graph4
            if is_dual:
                data["combo_dual"] = new_line
                data["login_client_id"] = parts[4]
                data["login_refresh_token"] = parts[5]
            now = datetime.now().isoformat()
            data["updated_at"] = now
            data["last_alive_at"] = now
            fp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass
    # accounts.txt（四段）
    if COMBO_FILE.exists():
        changed = False
        kept: list[str] = []
        for l in COMBO_FILE.read_text(encoding="utf-8").splitlines():
            e = l.split("----")[0].strip() if "----" in l else ""
            if e == email and not l.strip().startswith("#"):
                kept.append(graph4)
                changed = True
            else:
                kept.append(l)
        if changed:
            COMBO_FILE.write_text("\n".join(kept) + "\n", encoding="utf-8")
    # accounts_dual.txt（六段）
    if is_dual and DUAL_FILE.exists():
        changed = False
        kept = []
        for l in DUAL_FILE.read_text(encoding="utf-8").splitlines():
            e = l.split("----")[0].strip() if "----" in l else ""
            if e == email and not l.strip().startswith("#"):
                kept.append(new_line)
                changed = True
            else:
                kept.append(l)
        if changed:
            DUAL_FILE.write_text("\n".join(kept) + "\n", encoding="utf-8")


@app.post("/api/rescue")
def rescue_accounts(req: RescueRequest) -> JSONResponse:
    """对选中账号跑 scripts.rescue_login，回写 token / 恢复邮箱并累计 rescue_count。"""
    if not RESCUE_READY or rescue_and_persist is None or rescue_proxy_raw is None:
        return JSONResponse({
            "ok": False,
            "implemented": False,
            "message": "救援脚本 scripts.rescue_login 无法导入，暂不可用。",
            "results": [],
        })
    rows = _load_accounts()
    want = set(req.emails) if req.emails else None
    tasks: list[dict[str, Any]] = []
    for r in rows:
        if want is not None and r["email"] not in want:
            continue
        if not (r.get("password") or "").strip():
            continue
        tasks.append(r)
    if not tasks:
        return JSONResponse({
            "ok": True,
            "implemented": True,
            "total": 0,
            "ok_count": 0,
            "results": [],
            "message": "无可救援账号（需有密码）。",
        })

    proxy = rescue_proxy_raw((req.proxy or "").strip())
    use_pool = bool(req.use_proxy_pool)
    if (req.proxy or "").strip():
        proxy_pool.ensure_templates([(req.proxy or "").strip()], provider="web")
    if use_pool and not proxy_pool.pool_stats().get("enabled") and not proxy:
        return JSONResponse({
            "ok": False,
            "implemented": True,
            "message": "代理池为空，请先在「代理池」页添加或在重登时填写代理。",
            "results": [],
        })
    conc = max(1, min(int(req.concurrency or 1), 2, len(tasks)))
    results: list[dict[str, Any]] = []

    def work(row: dict[str, Any]) -> dict[str, Any]:
        email = row["email"]
        one_proxy = proxy
        proxy_meta: dict[str, Any] = {}
        if use_pool:
            one_proxy, proxy_meta = proxy_pool.resolve_for_email(email, fallback=proxy)
            one_proxy = rescue_proxy_raw(one_proxy or proxy)
        try:
            out = rescue_and_persist(
                email,
                row.get("password") or "",
                one_proxy,
                recovery_email=row.get("recovery_email") or "",
                write=True,
                accounts_dir=ACCOUNTS_DIR,
            )
        except Exception as exc:  # noqa: BLE001
            logger.exception("救援异常 %s: %s", email, exc)
            out = {"ok": False, "email": email, "reason": f"{type(exc).__name__}: {exc}"}
        out.setdefault("email", email)
        if use_pool and proxy_meta.get("proxy_id"):
            proxy_pool.record_result(
                proxy_meta["proxy_id"],
                success=bool(out.get("ok")),
                reg_country="",
                purpose="rescue",
                email=email,
                error=(out.get("reason") or "") if not out.get("ok") else "",
            )
        if out.get("ok") and out.get("refresh_token"):
            try:
                v = _verify_one(email, out["refresh_token"], "", False)
                _cache_verify(v)
                out["verify_ok"] = bool(v.get("ok"))
                out["verify_summary"] = v.get("summary", "")
            except Exception as exc:  # noqa: BLE001
                out["verify_ok"] = False
                out["verify_summary"] = str(exc)[:120]
        slim = {k: v for k, v in out.items() if k != "refresh_token" or not out.get("ok")}
        return slim

    with ThreadPoolExecutor(max_workers=conc) as ex:
        for f in as_completed([ex.submit(work, t) for t in tasks]):
            results.append(f.result())

    ok_n = sum(1 for r in results if r.get("ok"))
    return JSONResponse({
        "ok": True,
        "implemented": True,
        "total": len(results),
        "ok_count": ok_n,
        "results": results,
    })


@app.post("/api/keepalive")
def keepalive(req: KeepaliveRequest) -> JSONResponse:
    """对选中（或全部）账号并发跑 keepalive_one：refresh→access→GET /me+列信→轮换回写。"""
    if not KEEPALIVE_READY or keepalive_one is None:
        return JSONResponse({"ok": False, "implemented": False,
                             "message": "保活脚本 scripts.keepalive 无法导入，暂不可用。"})
    proxy_url = _proxy_url(req.proxy)
    rows = _load_accounts()
    want = set(req.emails) if req.emails else None
    tasks: list[tuple[str, str]] = []  # (email, combo_line)
    for r in rows:
        if want is not None and r["email"] not in want:
            continue
        line = r.get("combo_dual") or r.get("combo") or ""
        if not line or not r.get("has_token"):
            continue
        tasks.append((r["email"], line))
    if not tasks:
        return JSONResponse({"ok": True, "implemented": True, "results": [],
                             "message": "无可保活账号（缺 refresh_token）。"})

    conc = max(1, min(int(req.concurrency or 5), 5, len(tasks)))
    results: list[dict[str, Any]] = []

    def work(item: tuple[str, str]) -> dict[str, Any]:
        email, line = item
        try:
            res = keepalive_one(line, proxy_url)
        except Exception as exc:  # noqa: BLE001
            return {"email": email, "ok": False, "detail": f"异常:{exc}"[:120]}
        res.setdefault("email", email)
        return res

    with ThreadPoolExecutor(max_workers=conc) as ex:
        for f in as_completed([ex.submit(work, t) for t in tasks]):
            results.append(f.result())

    # 回写轮换后的新 refresh_token（仅在真的轮换时写盘）
    with _save_lock:
        for r in results:
            if r.get("ok") and r.get("rotated") and r.get("new_line"):
                _writeback_keepalive(r.get("email", ""), r["new_line"])
    # 缓存保活结果到 meta（作为测活状态）
    for r in results:
        if r.get("skip"):
            continue
        email = r.get("email")
        if not email:
            continue
        now = datetime.now().isoformat()
        _update_meta(email, {"verify": {
            "ok": bool(r.get("ok")),
            "usable": ["graph"] if r.get("ok") else [],
            "graph": bool(r.get("profile")),
            "checked_at": now,
            "source": "keepalive",
        }})
        json_patch: dict[str, Any] = {"updated_at": now}
        if r.get("ok"):
            json_patch["last_alive_at"] = now
        _patch_account_json(email, json_patch)

    ok_n = sum(1 for r in results if r.get("ok"))
    rot_n = sum(1 for r in results if r.get("rotated"))
    return JSONResponse({
        "ok": True, "implemented": True, "total": len(results),
        "alive": ok_n, "rotated": rot_n,
        "results": [{
            "email": r.get("email"),
            "ok": bool(r.get("ok")),
            "profile": r.get("profile"),
            "message": r.get("message"),
            "rotated": bool(r.get("rotated")),
            "skip": bool(r.get("skip")),
            "detail": r.get("detail", ""),
        } for r in results],
    })


@app.post("/api/replenish")
def replenish_pool_api(req: ReplenishRequest) -> JSONResponse:
    """回补收码池：把选中/全部账号中 graph 可用的四段式去重追加进 proof pool。"""
    try:
        from outlook_api_reg.graph_mail import probe_token
        from outlook_api_reg.proof_pool import pool_path
    except Exception as exc:  # noqa: BLE001
        return JSONResponse({"ok": False, "implemented": False, "message": f"收码池模块不可用: {exc}"})
    pool = pool_path() or (PROJECT_DIR.parent / "1000outlook.txt")
    pool = Path(pool)
    proxy_url = _proxy_url(req.proxy)
    rows = _load_accounts()
    want = set(req.emails) if req.emails else None
    existing: set[str] = set()
    if pool.exists():
        for l in pool.read_text(encoding="utf-8", errors="replace").splitlines():
            l = l.strip()
            if l and not l.startswith("#") and "----" in l:
                existing.add(l.split("----")[0].lower())
    added = dup = bad = 0
    to_write: list[str] = []
    for r in rows:
        if want is not None and r["email"] not in want:
            continue
        parts = (r.get("combo") or "").split("----")
        if len(parts) < 4 or not parts[3]:
            bad += 1
            continue
        email, rt = parts[0], parts[3]
        four = "----".join(parts[:4])
        if email.lower() in existing:
            dup += 1
            continue
        if req.verify:
            try:
                pr = probe_token(email, rt, proxy_url=proxy_url)
            except Exception:  # noqa: BLE001
                pr = {"usable": []}
            if not pr.get("usable"):
                bad += 1
                continue
        to_write.append(four)
        existing.add(email.lower())
        added += 1
    if to_write:
        with _save_lock:
            pool.parent.mkdir(parents=True, exist_ok=True)
            with pool.open("a", encoding="utf-8") as fp:
                for l in to_write:
                    fp.write(l + "\n")
    return JSONResponse({"ok": True, "implemented": True, "pool": str(pool),
                         "added": added, "duplicate": dup, "skipped": bad})


# ---------------------------------------------------------------------------
# 代理池
# ---------------------------------------------------------------------------


@app.get("/api/proxy-pool")
def get_proxy_pool(
    provider: Optional[str] = Query(None),
    limit: int = Query(5000, ge=1, le=20000),
    offset: int = Query(0, ge=0),
) -> JSONResponse:
    store = proxy_pool.load_store()
    return JSONResponse({
        "ok": True,
        "backend": proxy_pool.storage_backend(),
        "stats": proxy_pool.pool_stats(provider=provider),
        "providers": proxy_pool.list_providers(),
        "settings": store.get("settings") or {},
        "proxies": proxy_pool.list_entries(for_api=True, provider=provider, limit=limit, offset=offset),
        "bindings": proxy_pool.bindings_for_api(limit=min(limit, 2000), offset=offset),
        "file": str(proxy_pool.pool_file()),
    })


@app.get("/api/proxy-pool/analytics")
def get_proxy_pool_analytics(
    provider: Optional[str] = Query(None),
    country: Optional[str] = Query(None),
    reg_country: Optional[str] = Query(None),
    days: int = Query(0, ge=0, le=365),
) -> JSONResponse:
    return JSONResponse({
        "ok": True,
        "analytics": proxy_pool.proxy_analytics(
            provider=provider,
            country=country,
            reg_country=reg_country,
            days=days,
        ),
    })


@app.get("/api/proxy-pool/analytics/timeseries")
def get_proxy_pool_timeseries(
    provider: Optional[str] = Query(None),
    country: Optional[str] = Query(None),
    reg_country: Optional[str] = Query(None),
    days: int = Query(30, ge=1, le=365),
    group_by: str = Query("provider"),
) -> JSONResponse:
    return JSONResponse({
        "ok": True,
        "timeseries": proxy_pool.proxy_analytics_timeseries(
            provider=provider,
            country=country,
            reg_country=reg_country,
            days=days,
            group_by=group_by,
        ),
    })


@app.post("/api/proxy-pool/backfill-countries")
def backfill_proxy_countries(force: bool = Query(False)) -> JSONResponse:
    result = proxy_pool.backfill_proxy_countries(force=force)
    return JSONResponse({
        "ok": True,
        **result,
        "proxies": proxy_pool.list_entries(for_api=True),
    })


@app.post("/api/proxy-pool/ensure")
def ensure_proxy_pool(req: ProxyPoolEnsureRequest) -> JSONResponse:
    result = proxy_pool.ensure_templates(
        req.templates,
        text=req.text,
        provider=(req.provider or "web").strip(),
        country=(req.country or "").strip(),
    )
    return JSONResponse({
        "ok": True,
        **result,
        "stats": proxy_pool.pool_stats(),
        "proxies": proxy_pool.list_entries(for_api=True),
    })


@app.post("/api/proxy-pool")
def add_proxy_pool(req: ProxyPoolAddRequest) -> JSONResponse:
    templates = list(req.templates or [])
    if req.text:
        for line in req.text.replace("\r", "").split("\n"):
            line = line.strip()
            if line and not line.startswith("#"):
                templates.append(line)
    if not templates:
        raise HTTPException(status_code=400, detail="请提供至少一条代理模板。")
    created = proxy_pool.add_proxies(
        templates,
        label=(req.label or "").strip(),
        provider=(req.provider or "").strip(),
        country=(req.country or "").strip(),
    )
    return JSONResponse({
        "ok": True,
        "added": len(created),
        "proxies": proxy_pool.list_entries(for_api=True),
        "stats": proxy_pool.pool_stats(),
    })


@app.put("/api/proxy-pool/{proxy_id}")
def update_proxy_pool_item(proxy_id: str, req: ProxyPoolUpdateRequest) -> JSONResponse:
    ent = proxy_pool.update_proxy(
        proxy_id,
        label=req.label,
        template=req.template,
        provider=req.provider,
        country=req.country,
        enabled=req.enabled,
    )
    if not ent:
        raise HTTPException(status_code=404, detail="代理不存在。")
    return JSONResponse({"ok": True, "proxy": {**ent, "template_masked": proxy_pool.mask_template(ent.get("template") or "")}})


@app.post("/api/proxy-pool/delete")
def delete_proxy_pool(req: ProxyPoolDeleteRequest) -> JSONResponse:
    n = proxy_pool.delete_proxies(req.ids or [])
    return JSONResponse({"ok": True, "deleted": n, "stats": proxy_pool.pool_stats()})


@app.post("/api/proxy-pool/check")
def check_proxy_pool(req: ProxyPoolCheckRequest) -> JSONResponse:
    results = proxy_pool.check_proxies(req.ids, timeout=max(5, min(int(req.timeout or 15), 60)))
    ok_n = sum(1 for r in results if r.get("ok"))
    return JSONResponse({
        "ok": True,
        "checked": len(results),
        "healthy": ok_n,
        "results": results,
        "stats": proxy_pool.pool_stats(),
        "proxies": proxy_pool.list_entries(for_api=True),
    })


@app.post("/api/proxy-pool/settings")
def update_proxy_pool_settings(req: ProxyPoolSettingsRequest) -> JSONResponse:
    settings = proxy_pool.update_settings(
        strategy=req.strategy,
        require_healthy=req.require_healthy,
        sticky_per_account=req.sticky_per_account,
    )
    return JSONResponse({"ok": True, "settings": settings})


@app.post("/api/proxy-pool/bind")
def bind_proxy_pool(req: ProxyPoolBindRequest) -> JSONResponse:
    store = proxy_pool.load_store()
    ent = proxy_pool.entry_by_id(store, req.proxy_id)
    if not ent:
        raise HTTPException(status_code=404, detail="代理不存在。")
    resolved = proxy_pool.resolve_template(ent.get("template") or "")
    if not resolved:
        raise HTTPException(status_code=400, detail="代理模板无效。")
    proxy_pool.bind_account(req.email.strip().lower(), req.proxy_id, resolved, purpose="manual")
    return JSONResponse({"ok": True, "email": req.email.strip().lower(), "resolved_masked": proxy_pool.mask_template(resolved)})


@app.post("/api/proxy-pool/unbind")
def unbind_proxy_pool(req: ProxyPoolUnbindRequest) -> JSONResponse:
    n = proxy_pool.unbind_accounts(req.emails or [])
    return JSONResponse({"ok": True, "removed": n})


@app.get("/api/health")
def health() -> JSONResponse:
    db_st = app_db.db_status()
    return JSONResponse(
        {
            "ok": True,
            "accounts_dir": str(ACCOUNTS_DIR),
            "database": db_st,
            "batch_ready": BATCH_READY,
            "dual_ready": DUAL_READY,
            "keepalive_ready": KEEPALIVE_READY,
            "rescue_ready": RESCUE_READY,
        }
    )


@app.get("/api/database")
def database_status() -> JSONResponse:
    return JSONResponse({"ok": True, **app_db.db_status()})


@app.post("/api/database/backup")
def database_backup() -> JSONResponse:
    try:
        dest = app_db.backup_database(tag="manual")
        return JSONResponse({"ok": True, "path": str(dest), "status": app_db.db_status()})
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/database/migrate")
def database_migrate() -> JSONResponse:
    """手动触发遗留 JSON 迁移（通常启动时已自动执行）。"""
    stats = app_db.migrate_legacy_files(ACCOUNTS_DIR)
    return JSONResponse({"ok": True, "migrated": stats, "status": app_db.db_status()})
