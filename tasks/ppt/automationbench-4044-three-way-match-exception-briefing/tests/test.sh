#!/usr/bin/env bash
set -eo pipefail
root="$(cd "$(dirname "$0")/.." && pwd)"; result="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"; reward="${REWARD_PATH:-/logs/verifier/reward.txt}"; mkdir -p "$(dirname "$result")" "$(dirname "$reward")"; exec "${PYTHON_BIN:-python3}" "$root/tests/judge.py" --artifact-dir "${1:-${ARTIFACT_DIR:-/app/output}}" --result "$result" --reward-file "$reward"
