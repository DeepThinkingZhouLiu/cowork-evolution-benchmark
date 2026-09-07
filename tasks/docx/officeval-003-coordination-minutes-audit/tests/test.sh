#!/bin/sh
set -eu
python /tests/judge.py --artifact-dir /app/output --result /logs/verifier/result.json --reward-file /logs/verifier/reward.txt
