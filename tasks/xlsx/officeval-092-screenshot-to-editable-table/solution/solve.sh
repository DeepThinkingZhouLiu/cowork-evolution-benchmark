#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
input_dir="${INPUT_DIR:-${task_root}/data/input_files}"
output_dir="${1:-${OUTPUT_DIR:-/app/output}}"
python_bin="${PYTHON_BIN:-python3}"

exec "${python_bin}" "${task_root}/solution/solve.py" \
  --input-dir "${input_dir}" \
  --output-dir "${output_dir}"
