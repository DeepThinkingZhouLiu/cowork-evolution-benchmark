#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
input_dir="${INPUT_DIR:-/app/input_files}"
output_dir="${OUTPUT_DIR:-/app/output}"

exec "${PYTHON_BIN:-python3}" "${task_root}/solution/solve.py" \
  --input-dir "${input_dir}" \
  --output-dir "${output_dir}"
