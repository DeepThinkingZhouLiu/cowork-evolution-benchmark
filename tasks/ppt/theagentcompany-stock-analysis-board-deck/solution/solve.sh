#!/usr/bin/env bash
set -eo pipefail
root="$(cd "$(dirname "$0")" && pwd)"
exec "${PYTHON_BIN:-python3}" "$root/solve.py" --input-dir "${INPUT_DIR:-/app/input_files}" --output-dir "${OUTPUT_DIR:-/app/output}" --golden-dir "${GOLDEN_DIR:-$root/Golden}"
