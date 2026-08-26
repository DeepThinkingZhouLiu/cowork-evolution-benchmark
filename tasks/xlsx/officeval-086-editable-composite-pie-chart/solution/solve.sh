#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
input_dir="${1:-${INPUT_DIR:-${task_root}/data/input_files}}"
output_dir="${2:-${OUTPUT_DIR:-/app/output}}"
python_bin="${PYTHON_BIN:-python3}"

mkdir -p "${output_dir}"
"${python_bin}" "${task_root}/solution/solve.py" \
  --input "${input_dir}/数据1.xlsx" \
  --output-dir "${output_dir}"
