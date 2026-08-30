#!/usr/bin/env python3
"""PPTArena Case 25 Oracle：回放逐单元格审计后的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "FixTableThemeColor_TestA.pptx"
OUTPUT_NAME = "FixTableThemeColor_GroundTruthA.pptx"
INPUT_SHA256 = "d0b3c9d5b9988a375252315d3300b3237f25684af23ed77a4330aa8021d10c30"
GOLDEN_SHA256 = "fa3fbef6d9d97a8a5f71f46dff48ef0399555050ecddb6555ba3bd7face53acb"


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
        raise ValueError("已逐单元格重建的本地 Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
