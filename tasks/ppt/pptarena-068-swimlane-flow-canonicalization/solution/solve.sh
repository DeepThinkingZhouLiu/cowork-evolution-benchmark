#!/usr/bin/env bash
set -eo pipefail
solution_root="$(cd "$(dirname "$0")" && pwd)"
python_bin="$PYTHON_BIN"
input_dir="$INPUT_DIR"
output_dir="$OUTPUT_DIR"
golden_dir="$GOLDEN_DIR"
test -n "$python_bin" || python_bin=python3
test -n "$input_dir" || input_dir=/app/input_files
test -n "$output_dir" || output_dir=/app/output
test -n "$golden_dir" || golden_dir="$solution_root/Golden"
exec "$python_bin" "$solution_root/solve.py" --input-dir "$input_dir" --output-dir "$output_dir" --golden-dir "$golden_dir"
