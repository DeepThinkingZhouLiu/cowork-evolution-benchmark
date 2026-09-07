#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
exec "${PYTHON_BIN:-python3}" "$task_root/solution/solve.py" --input-dir "${INPUT_DIR:-$task_root/data/input_files}" --output-dir "${OUTPUT_DIR:-/app/output}"
