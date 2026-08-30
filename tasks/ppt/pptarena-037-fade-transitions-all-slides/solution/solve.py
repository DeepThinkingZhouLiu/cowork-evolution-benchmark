#!/usr/bin/env python3
"""PPTArena Case 37 Oracle：回放已逐页、OOXML 和视觉审计的上游人工 Ground Truth。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "AddTransitions_TestA.pptx"
OUTPUT_NAME = "AddTransitions_GroundTruthA.pptx"
INPUT_SHA256 = "a5557a252a8b8668a916a106b54c9935dd5d1eef2eb76d5f2fa4efa5bd4529b4"
GOLDEN_SHA256 = "c9b75e9063f87e8aac5c262c4b32a48d572119a130c50a0fd04e52108ff209db"


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
        raise ValueError("已逐页、OOXML 和视觉审计的人工 Golden 缺失或 SHA256 不匹配")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
