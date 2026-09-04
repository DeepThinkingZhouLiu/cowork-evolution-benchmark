#!/usr/bin/env bash
set -euo pipefail
script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
python3 "$script_dir/solve.py" --input-dir "${INPUT_DIR:-$script_dir/../data/input_files}" --output-dir "${OUTPUT_DIR:-/app/output}"
