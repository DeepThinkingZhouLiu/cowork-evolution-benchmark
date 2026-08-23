#!/usr/bin/env bash
set -euo pipefail

TASK_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CANDIDATE_DIR="${1:-${CANDIDATE_DIR:-/app/output}}"
python3 "${TASK_ROOT}/tests/judge.py"   --candidate-dir "$CANDIDATE_DIR"   --golden-dir "${TASK_ROOT}/solution"   --checks "${TASK_ROOT}/tests/checks.json"
