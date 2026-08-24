#!/usr/bin/env bash
set -euo pipefail

python /app/tests/judge.py \
  --artifact-dir /app/solution \
  --input-dir /app/data/input_files \
  --result /logs/artifacts/golden-judge-result.json

python /app/tests/negative_cases.py \
  --golden-dir /app/solution \
  --input-dir /app/data/input_files \
  --result /logs/artifacts/negative-cases.json
