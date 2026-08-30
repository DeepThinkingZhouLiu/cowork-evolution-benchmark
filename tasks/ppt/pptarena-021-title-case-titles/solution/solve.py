#!/usr/bin/env python3
"""PPTArena Case 21 Oracle：回放已逐页审计的上游人工 Ground Truth。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "FixTitleCase_TestA.pptx"
OUTPUT_NAME = "FixTitleCase_GroundTruthA.pptx"
INPUT_SHA256 = "5462ea413b6ecdf5ee13b0893c360e744886283949fed7a8b44fd78ed01e727f"
GOLDEN_SHA256 = "4b596d4f4fd84db6b331e7d06161643d7726e1c66979da393ecef780b65e1d89"


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
        raise ValueError("已逐页审计的人工 Golden 缺失或 SHA256 不匹配")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
