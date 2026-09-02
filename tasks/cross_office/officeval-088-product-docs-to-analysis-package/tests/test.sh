#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
python_bin="${PYTHON_BIN:-python3}"
result_path="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"
reward_path="${REWARD_PATH:-/logs/verifier/reward.txt}"
mkdir -p "$(dirname "${result_path}")" "$(dirname "${reward_path}")"
"${python_bin}" "${task_root}/tests/judge.py" --artifact-dir "${artifact_dir}" --result "${result_path}" || true
"${python_bin}" -c 'import json,sys,pathlib; p=pathlib.Path(sys.argv[1]); d=json.loads(p.read_text()) if p.is_file() else {}; pathlib.Path(sys.argv[2]).write_text("%.6f\n" % float(d.get("reward", d.get("score", 0.0))))' "${result_path}" "${reward_path}"
