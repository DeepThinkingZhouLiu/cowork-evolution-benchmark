#!/usr/bin/env python3
"""PPTArena Case 4 Oracle：回放已审计的上游人工 Ground Truth。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "FontSizePlus2Slide8TestA.pptx"
OUTPUT_NAME = "FontSizePlus2Slide8EditedA.pptx"
INPUT_SHA256 = "6beba374b051d77ba11636d826535fc3e391f245daf14d7547383a02798affd6"
GOLDEN_SHA256 = "0dcf73c3179e7a98438362eb96b2e2b84af00fedfea1c2e7071f36cf4d3f45f1"


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
        raise ValueError("已审计的人工 Golden 缺失或 SHA256 不匹配")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
