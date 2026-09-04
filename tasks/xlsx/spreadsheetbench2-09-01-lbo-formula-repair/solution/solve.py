#!/usr/bin/env python3
"""Harbor Golden 入口：从题目输入生成唯一修复工作簿。"""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
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
    with tempfile.TemporaryDirectory(prefix="cowork-evo-lbo-repair-") as temp_dir:
        subprocess.run(
            [
                sys.executable,
                str(source_processor),
                "--input",
                str((args.input_dir / "09_01_input.xlsx").resolve()),
                "--output",
                str(output_dir / "09_01_repaired.xlsx"),
                "--log",
                str(Path(temp_dir) / "repair_log.json"),
            ],
            check=True,
        )


if __name__ == "__main__":
    main()
