#!/usr/bin/env bash
set -euo pipefail
r="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)";exec "${PYTHON_BIN:-python3}" "$r/tests/judge.py" --artifact-dir "${ARTIFACT_DIR:-/app/output}" --result "${RESULT_FILE:-/logs/verifier/result.json}" --reward-file "${REWARD_FILE:-/logs/verifier/reward.txt}" --scoring-version v2
