#!/usr/bin/env python3
"""AutomationBench 1115 Oracle：回放逐题审计的 Golden。"""
from __future__ import annotations
import argparse, hashlib, shutil
from pathlib import Path

INPUTS = {
    "Competitive_Intelligence_Tracker.xlsx": "4bb463f0b76e21a875586f2f30ce666cde0a98280fa1d4ab281fa6d19c4f8883",
    "Analyst_Note_StrataCorp_v4.eml": "58854e7344764bef72b7e4b0ae1fd9313b3ad3308293792292da2d50f445bdee",
    "Pricing_Page_Change_Policy.eml": "36fd563f10d5a777d7c6ca90db270961011a3bab3481dea849ace4c69135dd7d",
    "Competitive_Intel_Slack.json": "f9fc7f315c4c4ba7fddf18dd4a64f8225846a45fa146921cad50938624d8b665",
    "AutomationBench_1115_Context.json": "6969d63a0d6163bbdaca24ac0495b73f3e0875997f54bf5143564fe3ab0102ed",
}
OUTPUT_NAME = "StrataCorp_v4_Competitive_Response_Brief.pptx"
GOLDEN_SHA256 = "f4ef3ab496a253a8d57e458beb3b42ec80f1ff69cfa998b0498fd21aee9ab49b"
def sha256(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def main():
    p=argparse.ArgumentParser(); p.add_argument("--input-dir",type=Path,required=True); p.add_argument("--output-dir",type=Path,required=True); p.add_argument("--golden-dir",type=Path,required=True); a=p.parse_args()
    for name, expected in INPUTS.items():
        path=a.input_dir/name
        if not path.is_file() or sha256(path)!=expected: raise ValueError(f"冻结输入缺失或 SHA256 不匹配：{name}")
    golden=a.golden_dir/OUTPUT_NAME
    if not golden.is_file() or sha256(golden)!=GOLDEN_SHA256: raise ValueError("已审计 Golden 缺失或 SHA256 不匹配")
    a.output_dir.mkdir(parents=True,exist_ok=True); shutil.copy2(golden,a.output_dir/OUTPUT_NAME); print(a.output_dir/OUTPUT_NAME)
if __name__=="__main__": main()
