#!/usr/bin/env bash
set -euo pipefail

input_dir="${INPUT_DIR:-/app/data/input_files}"
output_dir="${OUTPUT_DIR:-/app/output}"
exec "${PYTHON_BIN:-python3}" "$(dirname "${BASH_SOURCE[0]}")/solve.py" \
  --input-dir "$input_dir" --output-dir "$output_dir"
