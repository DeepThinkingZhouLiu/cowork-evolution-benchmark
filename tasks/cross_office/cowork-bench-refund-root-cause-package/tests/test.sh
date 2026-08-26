#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
exec "${PYTHON_BIN:-python3}" "${task_root}/tests/judge.py" --artifact-dir "${1:-${ARTIFACT_DIR:-/app/output}}"
