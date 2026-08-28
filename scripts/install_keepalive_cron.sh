#!/usr/bin/env bash
# 安装 crontab：每 2 天跑一次保活（凌晨 3:17）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CRON_SH="$ROOT/scripts/cron_keepalive.sh"
LOG="$ROOT/accounts/keepalive.log"
chmod +x "$CRON_SH" "$ROOT/scripts/keepalive_from_db.py" 2>/dev/null || true

MARKER="# outlook-auto-register keepalive_from_db"
ENTRY="17 3 */2 * * cd $ROOT && $CRON_SH >> $LOG 2>&1 $MARKER"

tmpdir=$(mktemp)
crontab -l 2>/dev/null | grep -v "keepalive_from_db" | grep -v "$MARKER" >"$tmpdir" || true
echo "$ENTRY" >>"$tmpdir"
crontab "$tmpdir"
rm -f "$tmpdir"
echo "已安装 crontab（每 2 天 03:17）:"
echo "  $ENTRY"
crontab -l | grep -F "keepalive_from_db" || true
