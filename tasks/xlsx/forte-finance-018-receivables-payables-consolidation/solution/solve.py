#!/usr/bin/env python3
"""从空输出目录回放 Finance-018 任务。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))

from reconcile_workbooks import generate


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    generate(args.input_dir, args.output_dir, Path(__file__).with_name("reconcile_workbooks.py"))


if __name__ == "__main__":
    main()
