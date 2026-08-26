#!/usr/bin/env python3
"""Harbor Golden 入口：复制可回放修复器并从题目输入生成交付物。"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    task_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=task_root / "data" / "input_files")
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    source_processor = Path(__file__).with_name("repair_workbook.py")
    output_processor = output_dir / "repair_workbook.py"
    shutil.copy2(source_processor, output_processor)
    subprocess.run(
        [
            sys.executable,
            str(output_processor),
            "--input",
            str((args.input_dir / "09_01_input.xlsx").resolve()),
            "--output",
            str(output_dir / "09_01_repaired.xlsx"),
            "--log",
            str(output_dir / "repair_log.json"),
        ],
        check=True,
    )


if __name__ == "__main__":
    main()
