#!/usr/bin/env python3
"""从 SQLite 账号库批量保活：跳过孵化期账号，轮换 refresh_token 写回数据库。

用法：
  python3 scripts/keepalive_from_db.py [--proxy h:p:u:p] [--concurrency 5] [--limit 0]

日志追加到 accounts/keepalive.log
"""
from __future__ import annotations

import argparse
import logging
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from outlook_api_reg import account_store  # noqa: E402
from outlook_api_reg import database as app_db  # noqa: E402
from outlook_api_reg import lifecycle  # noqa: E402
from scripts.keepalive import keepalive_one, _proxy_url  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
ACCOUNTS_DIR = ROOT / "accounts"
LOG_PATH = ACCOUNTS_DIR / "keepalive.log"


def _setup_log() -> logging.Logger:
    ACCOUNTS_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("keepalive_from_db")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(sh)
    return logger


def _combo_line(row: dict) -> str:
    return (row.get("combo_recovery") or row.get("combo_dual") or row.get("combo") or "").strip()


def _writeback(email: str, new_line: str) -> None:
    parts = new_line.split("----")
    if len(parts) < 4:
        return
    patch: dict = {
        "refresh_token": parts[3],
        "combo": "----".join(parts[:4]),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "last_alive_at": datetime.now().isoformat(timespec="seconds"),
    }
    if len(parts) >= 6 and parts[4] and parts[5]:
        # recovery 六段
        if "@" in parts[4]:
            patch["recovery_email"] = parts[4]
            patch["recovery_password"] = parts[5]
            patch["combo_recovery"] = new_line
        else:
            patch["login_client_id"] = parts[4]
            patch["login_refresh_token"] = parts[5]
            patch["combo_dual"] = new_line
    account_store.patch_account(email, patch)


def main() -> int:
    ap = argparse.ArgumentParser(description="SQLite 账号库保活（跳过孵化期）")
    ap.add_argument("--proxy", default="")
    ap.add_argument("--concurrency", type=int, default=5)
    ap.add_argument("--limit", type=int, default=0, help="最多处理 N 个（0=全部）")
    args = ap.parse_args()

    log = _setup_log()
    app_db.ensure_initialized(ACCOUNTS_DIR)
    proxy_url = _proxy_url(args.proxy)
    rows = account_store.list_accounts()

    tasks: list[tuple[str, str]] = []
    skipped = 0
    for r in rows:
        if not r.get("has_token") and not r.get("refresh_token"):
            continue
        if r.get("incubating") or lifecycle.is_incubating(r.get("created_at")):
            skipped += 1
            continue
        line = _combo_line(r)
        if not line:
            continue
        tasks.append((r["email"], line))

    if args.limit and args.limit > 0:
        tasks = tasks[: args.limit]

    log.info(
        "开始保活: 候选 %d · 孵化跳过 %d · 并发 %d",
        len(tasks),
        skipped,
        args.concurrency,
    )
    if not tasks:
        log.info("无可保活账号，结束")
        return 0

    ok = dead = rotated = 0

    def work(item: tuple[str, str]) -> dict:
        email, line = item
        try:
            res = keepalive_one(line, proxy_url)
        except Exception as exc:  # noqa: BLE001
            return {"email": email, "ok": False, "detail": str(exc)[:120]}
        res.setdefault("email", email)
        return res

    with ThreadPoolExecutor(max_workers=max(1, args.concurrency)) as pool:
        futs = [pool.submit(work, t) for t in tasks]
        for fut in as_completed(futs):
            res = fut.result()
            email = res.get("email") or "?"
            if res.get("ok"):
                ok += 1
                if res.get("rotated") and res.get("new_line"):
                    _writeback(email, res["new_line"])
                    rotated += 1
                log.info("[OK] %s rotated=%s", email, bool(res.get("rotated")))
            else:
                dead += 1
                log.warning("[DEAD] %s %s", email, res.get("detail", ""))

    log.info("保活完成: 存活 %d / 失效 %d / 轮换 %d / 孵化跳过 %d", ok, dead, rotated, skipped)
    return 0


if __name__ == "__main__":
    sys.exit(main())
