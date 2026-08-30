#!/usr/bin/env python3
"""PPTArena Case 67 Oracle：回放已审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "AccessibilityMasterCleanup_TestA.pptx"
OUTPUT_NAME = "AccessibilityMasterCleanup_GroundTruthA.pptx"
INPUT_SHA256 = "8e3ecec6a46f3ef3a41d764d4853359d919c840baaf5834a15e2c0814dc7e35b"
GOLDEN_SHA256 = "78fdf673fddd833cd5b0e70f4ac2605c05693dd91b11254471b0bec43c729b6e"


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
