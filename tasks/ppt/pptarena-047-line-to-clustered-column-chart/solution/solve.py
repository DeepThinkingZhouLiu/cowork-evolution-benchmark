#!/usr/bin/env python3
"""PPTArena Case 47 Oracle：回放已逐对象审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "ChangeLineGraphtoBar_TestA.pptx"
OUTPUT_NAME = "ChangeLineGraphtoBar_GroundTruthA.pptx"
INPUT_SHA256 = "badd9134a75308729132505dc232413229bf20914fa0fddb42f4bf186b420114"
GOLDEN_SHA256 = "6ae3c3b46082ec60ac053015c0e7f3761531a0ba866c00a2d82015f480e30873"


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
        raise ValueError("已审计 Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
