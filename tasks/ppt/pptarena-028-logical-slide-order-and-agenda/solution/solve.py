#!/usr/bin/env python3
"""PPTArena Case 28 Oracle：回放已逐页审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "FixAgendaAndOrder_TestA.pptx"
OUTPUT_NAME = "FixAgendaAndOrder_GroundTruthA.pptx"
INPUT_SHA256 = "bcbdf13b305029f1fca7c33d0b6b0020f84acf963bf066471d4be00971c2175b"
GOLDEN_SHA256 = "6079f8954c888ff2b10cc34e8bb6acfee0895b474edb10a4018739f4e018c554"


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
        raise ValueError("已逐页审计的 Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
