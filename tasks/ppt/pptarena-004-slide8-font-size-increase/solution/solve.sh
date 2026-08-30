#!/usr/bin/env bash
set -euo pipefail
solution_root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec "${PYTHON_BIN:-python3}" "${solution_root}/solve.py" \
  --input-dir "${INPUT_DIR:-/app/input_files}" \
  --output-dir "${OUTPUT_DIR:-/app/output}" \
  --golden-dir "${solution_root}/Golden"
