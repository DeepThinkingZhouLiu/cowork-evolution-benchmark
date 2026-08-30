#!/usr/bin/env python3
"""PPTArena Case 16 Oracle：回放从冻结输入逐页重建并审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "OrderPhotosTest.pptx"
OUTPUT_NAME = "OrderPhotosGroundTruth.pptx"
INPUT_SHA256 = "c17f8b37992c462006bf62a151b3f346df2dade7253f97027f69d715ab7f657c"
GOLDEN_SHA256 = "5791250ffc46eaa7a95b388d86c246b09030ce476b430b84a170d7c496c97f05"


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
        raise ValueError("已逐页重建的本地 Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
