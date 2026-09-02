#!/usr/bin/env bash
set -euo pipefail
r="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"; i="${INPUT_DIR:-/app/input_files}"; if [[ ! -d "$i" ]]; then i="$r/data/input_files"; fi
exec "${PYTHON_BIN:-python3}" "$r/solution/solve.py" --input-dir "$i" --output-dir "${OUTPUT_DIR:-/app/output}"
