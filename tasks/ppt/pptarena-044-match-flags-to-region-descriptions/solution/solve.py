#!/usr/bin/env python3
"""PPTArena Case 44 Oracle：回放已逐页审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "MatchFlagtoText_TestA.pptx"
OUTPUT_NAME = "MatchFlagtoText_GroundTruthA.pptx"
INPUT_SHA256 = "8c2c6f409684507f2d56cc205c8a6b17d69e65d5036748dbe6d2098e24f325e2"
GOLDEN_SHA256 = "2b34894adfd13a72970accb30d212b3239e6672095a7db079239dd9f53b09067"


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
