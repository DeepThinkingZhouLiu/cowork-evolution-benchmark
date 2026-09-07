#!/bin/sh
set -eu
python /solution/solve.py --input-dir "${INPUT_DIR:-/app/input_files}" --output-dir "${OUTPUT_DIR:-/app/output}"
