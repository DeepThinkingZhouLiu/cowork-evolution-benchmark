#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
result_path="${2:-${RESULT_PATH:-/logs/artifacts/judge-result.json}}"
if [[ ! -d "$(dirname "${result_path}")" ]]; then
  result_path="${TMPDIR:-/tmp}/cowork-bench-fy2026-judge-result.json"
fi
data_dir="${3:-${DATA_DIR:-${task_root}/data/input_files}}"
python_bin="${PYTHON_BIN:-python3}"

exec "${python_bin}" "${task_root}/tests/judge.py" \
  --artifact-dir "${artifact_dir}" \
  --data-dir "${data_dir}" \
  --result "${result_path}"
