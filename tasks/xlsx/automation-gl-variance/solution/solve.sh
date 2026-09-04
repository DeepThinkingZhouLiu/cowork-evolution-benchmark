#!/usr/bin/env bash
set -euo pipefail
python3 "$(dirname "$0")/solve.py" --input-dir "${INPUT_DIR:-/app/data/input_files}" --output-dir "${OUTPUT_DIR:-/app/output}"
