#!/usr/bin/env bash
set -eo pipefail
task_root="$(cd "$(dirname "$0")/.." && pwd)"
python_bin="$PYTHON_BIN"
result_path="$JUDGE_RESULT_PATH"
reward_path="$REWARD_PATH"
artifact_dir="$1"
test -n "$python_bin" || python_bin=python3
test -n "$result_path" || result_path=/logs/artifacts/judge-result.json
test -n "$reward_path" || reward_path=/logs/verifier/reward.txt
test -n "$artifact_dir" || artifact_dir="$ARTIFACT_DIR"
test -n "$artifact_dir" || artifact_dir=/app/output
mkdir -p "$(dirname "$result_path")" "$(dirname "$reward_path")"
exec "$python_bin" "$task_root/tests/judge.py" \
  --artifact-dir "$artifact_dir" \
  --result "$result_path" \
  --reward-file "$reward_path"
