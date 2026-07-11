#!/usr/bin/env bash
set -euo pipefail

PYTHON="/Users/Sue/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

if [ ! -x "$PYTHON" ]; then
  echo "找不到可用的 Python 运行环境：$PYTHON" >&2
  exit 1
fi

exec "$PYTHON" "$(dirname "$0")/scripts/portfolio.py" "$@"

