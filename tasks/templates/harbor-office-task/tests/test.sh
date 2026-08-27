#!/usr/bin/env bash
set -euo pipefail

task_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
artifact_dir="${1:-${ARTIFACT_DIR:-/app/output}}"
result_path="${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}"
reward_path="${REWARD_PATH:-/logs/verifier/reward.txt}"

mkdir -p "$(dirname "${result_path}")" "$(dirname "${reward_path}")"
# 制题人需要补充 judge.py，并保留以上 Harbor 路径契约。
echo "制题模板尚未实现 Judge。" >&2
exit 2
