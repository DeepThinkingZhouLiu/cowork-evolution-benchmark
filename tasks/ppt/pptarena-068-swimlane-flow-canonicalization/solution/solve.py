#!/usr/bin/env python3
"""PPTArena Case 68 Oracle：回放已审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "SwimlaneFlowCanonicalization_TestA.pptx"
OUTPUT_NAME = "SwimlaneFlowCanonicalization_GroundTruthA.pptx"
INPUT_SHA256 = "c5f0e852c29a8347cb29de3897d2c16cb11c84c5a9cd32dfec0e69444c756982"
GOLDEN_SHA256 = "ae5d4577d4368100d3b46dfb7256f5a75ea2044864dd0268560bfebf0d829b82"


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
