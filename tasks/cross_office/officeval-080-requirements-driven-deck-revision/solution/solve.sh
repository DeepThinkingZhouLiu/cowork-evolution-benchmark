#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
default_input="/app/input_files"
[[ -d "${default_input}" ]] || default_input="${task_root}/data/input_files"
input_dir="${1:-${INPUT_DIR:-${default_input}}}"
output_dir="${2:-${OUTPUT_DIR:-/app/output}}"

exec "${PYTHON_BIN:-python3}" "${task_root}/solution/solve.py" \
  --input-dir "${input_dir}" \
  --output-dir "${output_dir}"
