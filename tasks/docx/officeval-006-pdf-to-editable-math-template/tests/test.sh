#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
python_bin="${PYTHON_BIN:-python3}"

exec "${python_bin}" "${task_root}/tests/judge.py" \
  --artifact-dir "${artifact_dir}"
