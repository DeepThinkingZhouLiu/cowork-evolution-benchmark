#!/bin/sh
set -eu
root="$(cd "$(dirname "$0")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
result="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"
reward="${REWARD_PATH:-/logs/verifier/reward.txt}"
mkdir -p "$(dirname "$result")" "$(dirname "$reward")"
exec python3 "$root/tests/judge.py" --artifact-dir "$artifact_dir" --result "$result" --reward-file "$reward"
