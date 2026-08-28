#!/usr/bin/env python3
"""CLI：把外部 Outlook combo（本地文件或 qoderji email_inventory）并入本地账号库。

用法：
  # 从一份 4/6 段 combo 文本文件导入，标一下来源
  .venv/bin/python scripts/import_external_outlook.py --file accounts.txt --source manual

  # 从同机 qoderji 的 email_inventory 拉（默认路径探测 / QODERJI_EMAIL_DB）
  .venv/bin/python scripts/import_external_outlook.py --qoderji --dry-run
  .venv/bin/python scripts/import_external_outlook.py --qoderji --status consumed --status untried
  .venv/bin/python scripts/import_external_outlook.py --qoderji --qoderji-db /opt/qoderji/data/cards.db --limit 500
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from outlook_api_reg import external_import  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(
        description="导入外部 Outlook combo（自有账号资产合并，非新注册）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--file", help="本地 combo 文本文件路径（4/6 段，每行一条，兼容运维台粘贴导入格式）")
    ap.add_argument("--qoderji", action="store_true", help="从 qoderji 的 email_inventory 拉取")
    ap.add_argument(
        "--qoderji-db",
        help="qoderji sqlite 路径（可逗号分隔多个/带 glob）；默认取 QODERJI_EMAIL_DB 或 /opt/qoderji/{data/,}*.db",
    )
    ap.add_argument(
        "--status",
        action="append",
        dest="statuses",
        metavar="STATUS",
        help=f"qoderji 状态过滤，可重复出现（默认 {','.join(external_import.DEFAULT_QODERJI_STATUSES)}，"
        "dead=OAuth 永久失效，默认始终排除）",
    )
    ap.add_argument("--batch-id", help="只拉 qoderji 某一个 batch_id")
    ap.add_argument("--limit", type=int, help="最多拉取 / 导入多少条")
    ap.add_argument("--source", default="", help="写入 accounts.legacy_source / account_meta 的来源标记")
    ap.add_argument("--batch-label", default="", help="写入 batch_label，运维台按批次筛选用")
    ap.add_argument(
        "--no-skip-incubation",
        action="store_true",
        help="不跳过孵化期（默认外源号视为长效号，created_at 回填 30 天前）",
    )
    ap.add_argument("--dry-run", action="store_true", help="只统计、不写库")
    args = ap.parse_args()

    if not args.file and not args.qoderji:
        ap.error("必须指定 --file 或 --qoderji 之一")
    if args.file and args.qoderji:
        ap.error("--file 与 --qoderji 一次只能选一个")

    skip_incubation = not args.no_skip_incubation

    if args.file:
        path = Path(args.file)
        if not path.exists():
            print(f"文件不存在: {path}", file=sys.stderr)
            return 1
        text = path.read_text(encoding="utf-8", errors="replace")
        combos, parse_stats = external_import.parse_combo_text(text)
        if args.limit:
            combos = combos[: args.limit]
        result = external_import.import_combos(
            combos,
            source=args.source,
            batch_label=args.batch_label,
            skip_incubation=skip_incubation,
            dry_run=args.dry_run,
        )
        result["invalid"] += parse_stats.get("invalid", 0)
        result["parsed_from_file"] = len(combos)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    statuses = tuple(args.statuses) if args.statuses else external_import.DEFAULT_QODERJI_STATUSES
    result = external_import.import_from_qoderji(
        db_path=args.qoderji_db,
        statuses=statuses,
        batch_id=args.batch_id,
        limit=args.limit,
        batch_label=args.batch_label,
        skip_incubation=skip_incubation,
        dry_run=args.dry_run,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("ok", True) else 1


if __name__ == "__main__":
    raise SystemExit(main())
