#!/usr/bin/env bash
set -eo pipefail
solution_root="$(cd "$(dirname "$0")" && pwd)"
python_bin="${PYTHON_BIN:-python3}"
input_dir="${INPUT_DIR:-/app/input_files}"
output_dir="${OUTPUT_DIR:-/app/output}"
golden_dir="${GOLDEN_DIR:-$solution_root/Golden}"
exec "$python_bin" "$solution_root/solve.py" --input-dir "$input_dir" --output-dir "$output_dir" --golden-dir "$golden_dir"
