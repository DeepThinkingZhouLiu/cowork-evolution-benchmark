#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
input_dir="${INPUT_DIR:-/app/input_files}"
if [[ ! -d "${input_dir}" ]]; then
  input_dir="${task_root}/data/input_files"
fi
result_path="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"
reward_path="${REWARD_PATH:-/logs/verifier/reward.txt}"
mkdir -p "$(dirname "${result_path}")" "$(dirname "${reward_path}")"

"${PYTHON_BIN:-python3}" "${task_root}/tests/judge.py" \
  --artifact-dir "${artifact_dir}" \
  --input-dir "${input_dir}" \
  --result "${result_path}" || true
"${PYTHON_BIN:-python3}" -c 'import json,sys,pathlib; p=pathlib.Path(sys.argv[1]); d=json.loads(p.read_text()) if p.is_file() else {}; pathlib.Path(sys.argv[2]).write_text("%.6f\n" % float(d.get("reward", d.get("score", 0.0))))' "${result_path}" "${reward_path}"
