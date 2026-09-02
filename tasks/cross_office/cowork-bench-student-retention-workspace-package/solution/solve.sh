#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
default_input="/app/input_files"
[[ -d "${default_input}" ]] || default_input="${task_root}/data/input_files"
exec "${PYTHON_BIN:-python3}" "${task_root}/solution/solve.py" --input-dir "${1:-${INPUT_DIR:-${default_input}}}" --output-dir "${2:-${OUTPUT_DIR:-/app/output}}"
