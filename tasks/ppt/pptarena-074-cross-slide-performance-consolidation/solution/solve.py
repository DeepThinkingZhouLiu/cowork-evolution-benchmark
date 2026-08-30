#!/usr/bin/env python3
"""PPTArena Case 74 Oracle：回放已审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "CrossSlideDataConsolidation_TestA.pptx"
OUTPUT_NAME = "CrossSlideDataConsolidation_GroundTruthA.pptx"
INPUT_SHA256 = "76214e891a723b38a67f17012099427b66c201b066150a76e0de5ed7befe3c07"
GOLDEN_SHA256 = "c8cc0dc7f48bfa7b2bf95f2675070c34af0bb203a573b68fb69ff43e34a0a103"


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
