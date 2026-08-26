#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
input_dir="${1:-${INPUT_DIR:-/app/data/input_files}}"
output_dir="${2:-${OUTPUT_DIR:-/app/output}}"
workspace_dir="${3:-${WORKSPACE_DIR:-${output_dir}/.deck_workspace}}"
python_bin="${PYTHON_BIN:-python3}"

exec "${python_bin}" "${task_root}/solution/solve.py" \
  --input-dir "${input_dir}" \
  --output-dir "${output_dir}" \
  --workspace-dir "${workspace_dir}"
