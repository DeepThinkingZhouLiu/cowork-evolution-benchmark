#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${ARTIFACT_DIR:-/app/output}"
result_file="${RESULT_FILE:-/logs/verifier/result.json}"
reward_file="${REWARD_FILE:-/logs/verifier/reward.txt}"

exec "${PYTHON_BIN:-python3}" "${task_root}/tests/judge.py" \
  --artifact-dir "${artifact_dir}" \
  --result "${result_file}" \
  --reward-file "${reward_file}"
