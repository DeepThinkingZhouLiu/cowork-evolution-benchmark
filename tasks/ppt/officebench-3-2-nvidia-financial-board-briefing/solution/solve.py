#!/usr/bin/env python3
"""校验 OfficeBench 3-2 输入后复制人工审计 Golden。"""
from __future__ import annotations
import argparse, hashlib, shutil
from pathlib import Path

SOURCE_SHA256 = "33b606d43e02d2808822258ce6b4edfcd7c7bb7ed471f5395583286707a45626"
GOLDEN_SHA256 = "875cce0e383842abeacb62b201dcc8caa0776a89a9366595c1a083792272e4cd"
OUTPUT_NAME = "NVIDIA_Q1_FY2024_Management_Briefing.pptx"

def sha256(path: Path) -> str: return hashlib.sha256(path.read_bytes()).hexdigest()

def main() -> None:
    parser=argparse.ArgumentParser(); parser.add_argument("--input-dir",type=Path,required=True); parser.add_argument("--output-dir",type=Path,required=True); parser.add_argument("--golden-dir",type=Path,required=True); args=parser.parse_args()
    source=args.input_dir/"financial_report.pdf"; golden=args.golden_dir/OUTPUT_NAME
    if not source.is_file() or sha256(source)!=SOURCE_SHA256: raise SystemExit("冻结输入缺失或 SHA256 不匹配")
    if not golden.is_file() or sha256(golden)!=GOLDEN_SHA256: raise SystemExit("Golden 缺失或 SHA256 不匹配")
    args.output_dir.mkdir(parents=True,exist_ok=True); target=args.output_dir/OUTPUT_NAME; shutil.copy2(golden,target); print(target)

if __name__=="__main__": main()
