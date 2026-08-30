#!/usr/bin/env python3
"""PPTArena Case 27 Oracle：回放逐图审计后的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "FixImagePorportions_TestA.pptx"
OUTPUT_NAME = "FixImagePorportions_GroundTruthA.pptx"
INPUT_SHA256 = "026285dad008ed551128b04548909d8e4377324b67f8428cb6a68d2c87d9555f"
GOLDEN_SHA256 = "0d7ae976a9601c6a803d7b64220ea972fcb059be85dc4b263dafa04b97727206"


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
        raise ValueError("已逐图重建的本地 Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
