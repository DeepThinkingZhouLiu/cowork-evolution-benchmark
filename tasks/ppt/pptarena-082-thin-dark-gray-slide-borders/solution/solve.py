#!/usr/bin/env python3
"""PPTArena Case 82 Oracle：回放经逐页审计合格的上游人工 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "AddSlideBorderTestB_AccountantCVResume.pptx"
OUTPUT_NAME = "AddSlideBorderGroundTruthB_AccountantCVResume.pptx"
INPUT_SHA256 = "bff930968572cce3f83f3bbb3eb26889090edfb7aceb681b109958ea78216ff2"
GOLDEN_SHA256 = "b3cd1e8595deb342a8b2a7d2d3de323767d4754bb8f4037ac355334660be019b"


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
        raise ValueError("已逐页审计的上游人工 Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
