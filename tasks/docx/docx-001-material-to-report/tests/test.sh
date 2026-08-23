#!/usr/bin/env bash
set -euo pipefail
task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-$task_root/solution}}"
python3 "$task_root/tests/judge.py" "$artifact_dir"
