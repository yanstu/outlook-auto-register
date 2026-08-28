#!/usr/bin/env bash
# 定时保活入口：从 SQLite 读号、跳过孵化期、写 keepalive.log
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
mkdir -p accounts
export PYTHONUNBUFFERED=1
if [[ -x "$ROOT/.venv/bin/python" ]]; then
  PY="$ROOT/.venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  PY=python3
else
  echo "python3 not found" >&2
  exit 1
fi
exec "$PY" "$ROOT/scripts/keepalive_from_db.py" "$@"
