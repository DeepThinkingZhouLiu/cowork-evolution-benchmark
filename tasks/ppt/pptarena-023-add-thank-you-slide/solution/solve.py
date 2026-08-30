#!/usr/bin/env python3
"""PPTArena Case 23 Oracle：回放已逐页审计的上游人工 Ground Truth。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "AddThankYouSlide_TestA.pptx"
OUTPUT_NAME = "AddThankYouSlide_GroundTruthA.pptx"
INPUT_SHA256 = "cd6c301ffa1fe987814217e02220e63404c9481e20aebc9d2fb3c3394727ebbc"
GOLDEN_SHA256 = "ca53da109e42215fa2e45d85cf974b56dd0e107167062e676e0fbb51e539dc1e"


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
