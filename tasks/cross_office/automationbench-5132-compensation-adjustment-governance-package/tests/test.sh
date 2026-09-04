#!/usr/bin/env bash
set -euo pipefail
root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
exec "${PYTHON_BIN:-python3}" "$root/tests/judge.py" \
  --artifact-dir "${ARTIFACT_DIR:-/app/output}" \
  --result "${RESULT_FILE:-/logs/verifier/result.json}" \
  --reward-file "${REWARD_FILE:-/logs/verifier/reward.txt}" \
  --scoring-version "${SCORING_VERSION:-v2}"
