#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
if [[ -d /app/input_files && -d /app/output ]]; then
  default_input=/app/input_files
  default_output=/app/output
else
  default_input="${task_root}/data/input_files"
  default_output="${task_root}/solution/Golden"
fi

input_dir="${1:-${INPUT_DIR:-${default_input}}}"
output_dir="${2:-${OUTPUT_DIR:-${default_output}}}"

exec "${PYTHON_BIN:-python3}" "${task_root}/solution/solve.py" \
  --input-dir "${input_dir}" \
  --output-dir "${output_dir}"
