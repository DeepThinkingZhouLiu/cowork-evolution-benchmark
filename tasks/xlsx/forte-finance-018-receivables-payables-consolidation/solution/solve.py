#!/usr/bin/env python3
"""从三份输入生成唯一的跨期整合工作簿。"""
from __future__ import annotations
import argparse, shutil, tempfile
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parent))
from reconcile_workbooks import generate

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input-dir",type=Path,required=True); ap.add_argument("--output-dir",type=Path,required=True); args=ap.parse_args(); args.output_dir.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="cowork-evo-forte-") as temp:
        temp_path=Path(temp); generate(args.input_dir,temp_path); shutil.copy2(temp_path/"往来款跨期整合.xlsx",args.output_dir/"往来款跨期整合.xlsx")
if __name__=="__main__": main()
