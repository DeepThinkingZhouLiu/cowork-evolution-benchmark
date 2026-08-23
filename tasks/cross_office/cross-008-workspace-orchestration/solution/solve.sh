#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
output_dir="${1:-${OUTPUT_DIR:-/app/output}}"
mkdir -p "$output_dir"
python3 "$task_root/solution/solve.py" --input-dir "$task_root/data/input_files" --output-dir "$output_dir"
