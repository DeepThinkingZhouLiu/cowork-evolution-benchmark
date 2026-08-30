#!/usr/bin/env python3
"""PPTArena Case 41 Oracle：回放从冻结输入逐对象重建并通过视觉审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "MatchBackgroundtoPicture_TestA.pptx"
OUTPUT_NAME = "MatchBackgroundtoPicture_GroundTruthA.pptx"
INPUT_SHA256 = "8c789d83d7063549ab3a198a7363e7e6d9feda84ea20865ae00a6307d3be52cd"
GOLDEN_SHA256 = "ac8d5b8457fa8be04d58acd326201e66a21caf505210cc90eb82aa92fddae26d"


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
        raise ValueError("已重建并通过图片/文字/渲染审计的本地 Golden 缺失或 SHA256 不匹配")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
