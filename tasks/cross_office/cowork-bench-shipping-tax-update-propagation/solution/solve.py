#!/usr/bin/env python3
"""生成本题 Golden 交付物。"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from shipping_tax_update_processor import generate_package


def main() -> None:
    task_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=task_root / "data" / "input_files")
    parser.add_argument("--output-dir", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()

    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    source_processor = Path(__file__).with_name("shipping_tax_update_processor.py").resolve()
    target_processor = output_dir / source_processor.name
    if source_processor != target_processor.resolve():
        shutil.copy2(source_processor, target_processor)
    generate_package(args.input_dir.resolve(), output_dir)


if __name__ == "__main__":
    main()
