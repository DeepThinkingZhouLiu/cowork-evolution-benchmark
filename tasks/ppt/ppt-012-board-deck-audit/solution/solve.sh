#!/usr/bin/env bash
set -euo pipefail

TASK_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
INPUT_DIR="${INPUT_DIR:-${TASK_ROOT}/data/input_files}"
OUTPUT_DIR="${1:-${OUTPUT_DIR:-/app/output}}"
PYTHON_BIN="${PYTHON_BIN:-python3}"

exec "${PYTHON_BIN}" "${TASK_ROOT}/solution/solve.py" --input-dir "${INPUT_DIR}" --output-dir "${OUTPUT_DIR}"
