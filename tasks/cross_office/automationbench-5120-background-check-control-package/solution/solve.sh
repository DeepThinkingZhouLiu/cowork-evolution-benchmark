#!/bin/sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
INPUT_DIR=${1:-/app/input_files}
OUTPUT_DIR=${2:-/app/output}

python3 "$SCRIPT_DIR/solve.py" --input-dir "$INPUT_DIR" --output-dir "$OUTPUT_DIR"
