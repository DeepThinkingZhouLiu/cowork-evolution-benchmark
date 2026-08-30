#!/usr/bin/env python3
"""PPTArena Case 14 Oracle：回放已审计的上游人工 Ground Truth。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "BoldTitleTestA_Game_theory_4.pptx"
OUTPUT_NAME = "Bolded_BoldTitleTestA.pptx"
INPUT_SHA256 = "a8f00f2d2c814e5ef9c81a7757a72588ee911af67c7d120c5ca078d3a29395c1"
GOLDEN_SHA256 = "d2547d52b8237a46f41fad8c13a1e5b10322443def2f9c52d7611bee106c1947"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--golden-dir", type=Path, required=True)
    args = parser.parse_args()

    source = args.input_dir / INPUT_NAME
    golden = args.golden_dir / OUTPUT_NAME
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("冻结输入缺失或 SHA256 不匹配")
    if not golden.is_file() or sha256(golden) != GOLDEN_SHA256:
        raise ValueError("已审计的人工 Golden 缺失或 SHA256 不匹配")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
