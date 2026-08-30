#!/usr/bin/env python3
"""PPTArena Case 36 Oracle：回放已逐页审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "AddSpeakerNotes_TestA.pptx"
OUTPUT_NAME = "AddSpeakerNotes_GroundTruth.pptx"
INPUT_SHA256 = "a81cd08f0bdd05079ba350c1fa32b95db8111e504ae657613707e69abf8fd201"
GOLDEN_SHA256 = "314fe0f387211ee4874d0be6bb2f6d90d48611e8fe841ee9354d46fea743d03d"


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
