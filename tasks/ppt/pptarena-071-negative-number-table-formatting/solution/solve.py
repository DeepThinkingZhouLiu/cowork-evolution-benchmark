#!/usr/bin/env python3
"""PPTArena Case 71 Oracle：回放已审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "ConditionalTableFormatting_Negatives_TestA.pptx"
OUTPUT_NAME = "ConditionalTableFormatting_Negatives_GroundTruthA.pptx"
INPUT_SHA256 = "0b64628a3f1169dc6dcfca2ea0acca80e1cc6aeba7b9aa9bc456e8f36b9a305f"
GOLDEN_SHA256 = "a19d9fda1134653cb7fd2a12cf94ca2d719ccae04c26d431a0a8d112e5640228"


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
