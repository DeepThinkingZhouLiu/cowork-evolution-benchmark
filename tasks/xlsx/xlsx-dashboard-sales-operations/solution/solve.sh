#!/usr/bin/env bash
set -euo pipefail

TASK_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
INPUT_DIR="${INPUT_DIR:-${TASK_ROOT}/data/input_files}"
if [[ -n "${1:-}" ]]; then
  OUTPUT_DIR="$1"
elif [[ -d /app ]]; then
  OUTPUT_DIR="${OUTPUT_DIR:-/app/output}"
else
  OUTPUT_DIR="${OUTPUT_DIR:-${TASK_ROOT}/output}"
fi
mkdir -p "$OUTPUT_DIR"
python3 "${TASK_ROOT}/solution/solve.py" --input-dir "$INPUT_DIR" --output-dir "$OUTPUT_DIR"
