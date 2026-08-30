#!/usr/bin/env python3
"""PPTArena Case 30 Oracle：回放已逐页、数据和图表对象审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "LineGraph_TestA.pptx"
OUTPUT_NAME = "LineGraph_GroundTruthA.pptx"
INPUT_SHA256 = "72fb98cc839fc54dc4b07dcdb4bf7eb8c6552ac6549d00c0010e3ab021c1f3e6"
GOLDEN_SHA256 = "73e05b261a46e300a0d4754123d3a4414aac12283b6bf6cd65e3a6d29035eefc"


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
