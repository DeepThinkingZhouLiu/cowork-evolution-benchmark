#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
result_path="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"
reward_path="${REWARD_PATH:-/logs/verifier/reward.txt}"
python_bin="${PYTHON_BIN:-python3}"
default_input_dir="${task_root}/data/input_files"
[[ -d /app/input_files ]] && default_input_dir="/app/input_files"
input_dir="${INPUT_DIR:-${default_input_dir}}"
mkdir -p "$(dirname "${result_path}")" "$(dirname "${reward_path}")"

set +e
INPUT_DIR="${input_dir}" "${python_bin}" "${task_root}/tests/judge.py" \
  --artifact-dir "${artifact_dir}" \
  --result "${result_path}"
judge_status=$?
set -e

if [[ -f "${result_path}" ]]; then
  "${python_bin}" -c 'import json,sys; from pathlib import Path; payload=json.loads(Path(sys.argv[1]).read_text(encoding="utf-8")); score=payload.get("reward", payload.get("score", 0.0)); Path(sys.argv[2]).write_text(f"{float(score):.6f}\n", encoding="utf-8")' "${result_path}" "${reward_path}"
  exit 0
fi
exit "${judge_status}"
