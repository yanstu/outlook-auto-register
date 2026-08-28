#!/usr/bin/env python3
"""导出长存账号（combo_recovery）到 accounts/long_lived.txt。

用法：
  python3 scripts/export_long_lived.py --min-days 7
  python3 scripts/export_long_lived.py --min-days 7 --out accounts/long_lived.txt
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from outlook_api_reg import account_store  # noqa: E402
from outlook_api_reg import database as app_db  # noqa: E402
from outlook_api_reg import lifecycle  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
ACCOUNTS_DIR = ROOT / "accounts"
DEFAULT_OUT = ACCOUNTS_DIR / "long_lived.txt"


def main() -> int:
    ap = argparse.ArgumentParser(description="导出长存账号（recovery 六段）")
    ap.add_argument("--min-days", type=float, default=7.0, help="最短存活天数（默认 7）")
    ap.add_argument("--out", default=str(DEFAULT_OUT), help="输出文件")
    args = ap.parse_args()

    app_db.ensure_initialized(ACCOUNTS_DIR)
    rows = account_store.list_accounts()
    lines: list[str] = []
    for r in rows:
        if not lifecycle.is_long_lived(r, min_days=args.min_days):
            continue
        line = lifecycle.combo_recovery_line(r)
        if line:
            lines.append(line)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
    print(f"导出 {len(lines)} 条（min_days={args.min_days}）→ {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
