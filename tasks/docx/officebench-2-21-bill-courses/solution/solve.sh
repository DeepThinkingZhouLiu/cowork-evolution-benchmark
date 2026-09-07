#!/bin/sh
set -eu
TASK_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
"${PYTHON_BIN:-python3}" "$TASK_ROOT/solution/solve.py" \
  --input "${INPUT_DIR:-$TASK_ROOT/data/input_files}/Bill_course_lists.xlsx" \
  --output "${OUTPUT_DIR:-/app/output}/bill_course_history_audit.docx"
