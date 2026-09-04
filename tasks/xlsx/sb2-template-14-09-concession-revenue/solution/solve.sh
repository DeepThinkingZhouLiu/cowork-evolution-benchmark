#!/usr/bin/env bash
set -euo pipefail
exec "${PYTHON_BIN:-python3}" "$(dirname "$0")/solve.py" \
  --input-dir "${1:-${INPUT_DIR:-/app/data/input_files}}" \
  --output-dir "${2:-${OUTPUT_DIR:-/app/output}}"
