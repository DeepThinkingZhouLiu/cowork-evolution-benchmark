#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
exec "${PYTHON_BIN:-python3}" "${task_root}/solution/solve.py" \
  --input-dir "${1:-${INPUT_DIR:-/app/input_files}}" \
  --output-dir "${2:-${OUTPUT_DIR:-/app/output}}"
