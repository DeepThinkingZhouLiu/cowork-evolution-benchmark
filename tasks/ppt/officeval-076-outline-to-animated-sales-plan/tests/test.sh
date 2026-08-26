#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
result_path="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"
if [[ ! -d "$(dirname "${result_path}")" ]]; then
  result_path="${TMPDIR:-/tmp}/officeval-076-judge-result.json"
fi

exec "${PYTHON_BIN:-python3}" "${task_root}/tests/judge.py" \
  --artifact-dir "${artifact_dir}" \
  --result "${result_path}"
