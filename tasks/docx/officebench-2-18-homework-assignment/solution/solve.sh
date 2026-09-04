#!/bin/sh
set -eu
TASK_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
"${PYTHON_BIN:-python3}" "$TASK_ROOT/solution/solve.py" --input-dir "${INPUT_DIR:-$TASK_ROOT/data/input_files}" --output-dir "${OUTPUT_DIR:-/app/output}"
