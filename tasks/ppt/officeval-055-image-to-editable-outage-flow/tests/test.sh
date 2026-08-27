#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
result_path="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"
reward_path="${REWARD_PATH:-/logs/verifier/reward.txt}"

mkdir -p "$(dirname "${result_path}")" "$(dirname "${reward_path}")"
exec "${PYTHON_BIN:-python3}" "${task_root}/tests/judge.py" \
  --artifact-dir "${artifact_dir}" \
  --result "${result_path}" \
  --reward-file "${reward_path}"
