#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
PYTHON_BIN=${NOCOMPYTOOLS_PYTHON:-python3}

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
    echo "找不到 Python：$PYTHON_BIN。请安装 Python 3 或设置 NOCOMPYTOOLS_PYTHON。" >&2
    exit 127
fi

cd -- "$SCRIPT_DIR"
exec "$PYTHON_BIN" nocompytools_entry.py "$@"
