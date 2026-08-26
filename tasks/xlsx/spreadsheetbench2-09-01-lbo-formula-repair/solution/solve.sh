#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
output_dir="${1:-${ARTIFACT_DIR:-/app/output}}"

exec "${PYTHON_BIN:-python3}" "${task_root}/solution/solve.py" \
  --input-dir "${task_root}/data/input_files" \
  --output-dir "${output_dir}"
