#!/usr/bin/env bash
set -euo pipefail
root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"; artifact="${1:-${ARTIFACT_DIR:-/app/output}}"; result="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"; reward="${REWARD_PATH:-/logs/verifier/reward.txt}"; mkdir -p "$(dirname "$result")" "$(dirname "$reward")"; exec python3 "$root/tests/judge.py" --artifact-dir "$artifact" --result "$result" --reward-file "$reward"
