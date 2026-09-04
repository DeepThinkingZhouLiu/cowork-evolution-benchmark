#!/bin/sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
ARTIFACT_DIR=${ARTIFACT_DIR:-${1:-/app/output}}
RESULT_FILE=${RESULT_FILE:-${2:-/logs/verifier/result.json}}
REWARD_FILE=${REWARD_FILE:-${3:-/logs/verifier/reward.txt}}

exec "${PYTHON_BIN:-python3}" "$SCRIPT_DIR/judge.py" --artifact-dir "$ARTIFACT_DIR" --result "$RESULT_FILE" --reward-file "$REWARD_FILE" --scoring-version "${SCORING_VERSION:-v2}"
