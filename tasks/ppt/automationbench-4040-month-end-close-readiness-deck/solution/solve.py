#!/usr/bin/env python3
"""AutomationBench 4040 Oracle：回放逐题审计的本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import shutil
from pathlib import Path


INPUTS = {
    "January_2026_Close_Checklist.xlsx": "067ffd3a28b7608c0e0f7f00e34aa62c3c8a928981460c39e4e7e8cab1694b70",
    "AutomationBench_4040_Context.json": "76003f8de514e771e61103bb80f2e681d199ffd9bfda0980c076e7992ec5667b",
}
OUTPUT_NAME = "January_2026_Close_Readiness_Briefing.pptx"
GOLDEN_SHA256 = "22596ab04141ed2fbacefbafaa754116ab2b99856a8950d5a55fe2bd05fa6ba8"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--golden-dir", type=Path, required=True)
    args = parser.parse_args()
    for name, expected in INPUTS.items():
        path = args.input_dir / name
        if not path.is_file() or sha256(path) != expected:
            raise ValueError(f"冻结输入缺失或 SHA256 不匹配：{name}")
    golden = args.golden_dir / OUTPUT_NAME
    if not golden.is_file() or sha256(golden) != GOLDEN_SHA256:
        raise ValueError("已审计 Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / OUTPUT_NAME
    shutil.copy2(golden, output)
    print(output)


if __name__ == "__main__":
    main()
