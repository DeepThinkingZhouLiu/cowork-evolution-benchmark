#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
exec python3 "$task_root/solution/solve.py" --input-dir "${1:-${INPUT_DIR:-/app/data/input_files}}" --output-dir "${2:-${OUTPUT_DIR:-/app/output}}"
