#!/usr/bin/env python3
"""PPTArena Case 29 Oracle：回放已逐页和图表对象审计的上游人工 Ground Truth。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "PieChart_TestA.pptx"
OUTPUT_NAME = "PieChart_GroundTruthA.pptx"
INPUT_SHA256 = "07911f768a13823f9e8a4a426a7d2667a16f8283a4237f2883abe8b24527415e"
GOLDEN_SHA256 = "1bc36f7e0165538765a77b9d8ddd5df473afdca0babe95cda229daa1c286c36c"


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
        raise ValueError("已逐页和图表对象审计的人工 Golden 缺失或 SHA256 不匹配")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
