#!/usr/bin/env python3
"""PPTArena Case 1 Oracle：回放已逐页复核的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "TranslateToEnglishButKeepallFrenchTestA_Bonjour mes amis!.pptx"
OUTPUT_NAME = "TranslatetoEnglishButKeepallFrenchGroundTruthA.pptx"
INPUT_SHA256 = "166e95dcf7883aa0ed60768e184c225851161a908f82223a7f8443f45cdd32dd"
GOLDEN_SHA256 = "63cc22f03dbcd0e645e95b7de03e0a11812adc79120df9a4343604de5c73e2e7"


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
