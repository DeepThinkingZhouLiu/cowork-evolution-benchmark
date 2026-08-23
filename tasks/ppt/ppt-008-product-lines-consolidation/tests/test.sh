#!/usr/bin/env bash
set -euo pipefail

TASK_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ARTIFACT_DIR="${1:-${ARTIFACT_DIR:-/app/output}}"
RESULT_PATH="${2:-${RESULT_PATH:-${ARTIFACT_DIR}/judge-result.json}}"
PYTHON_BIN="${PYTHON_BIN:-python3}"

exec "${PYTHON_BIN}" "${TASK_ROOT}/tests/judge.py" --artifact-dir "${ARTIFACT_DIR}" --result "${RESULT_PATH}"
