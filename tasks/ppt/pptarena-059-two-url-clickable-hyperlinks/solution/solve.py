#!/usr/bin/env python3
"""PPTArena Case 59 Oracle：回放从冻结输入只写入两个 hyperlink 的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "ConvertHyperLink_Test.pptx"
OUTPUT_NAME = "ConvertHyperLink_GroundTruth.pptx"
INPUT_SHA256 = "33132900f1422268aec0a20ba9e6ba0c2be1e6b9a4baeb9f0887c371506792f5"
GOLDEN_SHA256 = "2e976bad97a869c58f0fe380bd04911a62a613c5d271c446e7e8cb35cd51a94e"


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
        raise ValueError("已审计的本地 Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
