#!/usr/bin/env bash
set -euo pipefail
r="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)";exec "${PYTHON_BIN:-python3}" "$r/tests/judge.py" --artifact-dir "${ARTIFACT_DIR:-${1:-/app/output}}" --result "${RESULT_FILE:-${JUDGE_RESULT_PATH:-/logs/verifier/result.json}}" --reward-file "${REWARD_FILE:-${REWARD_PATH:-/logs/verifier/reward.txt}}" --scoring-version "${SCORING_VERSION:-v2}"
