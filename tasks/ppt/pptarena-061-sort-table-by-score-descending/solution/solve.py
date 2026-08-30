#!/usr/bin/env python3
"""PPTArena Case 61 Oracle：回放已逐对象审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUT_NAME = "SortByScoreAndCropImage169_Test.pptx"
OUTPUT_NAME = "SortByScoreAndCropImage169_GroundTruth.pptx"
INPUT_SHA256 = "dd2ee0141319451ea8b33b20c06b49e2985932f302ba6de1349566a864f7995c"
GOLDEN_SHA256 = "1c89355c12156d06eb4cd71d81dfc444767ed8dcee764b2a7078aa22ac5c16cf"


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
