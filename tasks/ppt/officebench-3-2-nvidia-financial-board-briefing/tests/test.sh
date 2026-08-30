#!/usr/bin/env bash
set -eo pipefail
task_root="$(cd "$(dirname "$0")/.." && pwd)"
python_bin="${PYTHON_BIN:-python3}"
result_path="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"
reward_path="${REWARD_PATH:-/logs/verifier/reward.txt}"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
mkdir -p "$(dirname "$result_path")" "$(dirname "$reward_path")"
exec "$python_bin" "$task_root/tests/judge.py" --artifact-dir "$artifact_dir" --result "$result_path" --reward-file "$reward_path"
