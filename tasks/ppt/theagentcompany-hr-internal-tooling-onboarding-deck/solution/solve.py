#!/usr/bin/env python3
from __future__ import annotations
import argparse,hashlib,shutil
from pathlib import Path
HASHES={"TAC_overview.md":"504f19054be9b3c2ed74841332dee564e88ee98c0f990d54c50b057605350e53","Personnell_File.odt":"e19e2ff3988ec2ed915f9e87571df40c3f7a27be2350e9c98a6263d65afc5025","Internal_Tooling_Source_Snapshot.json":"f6d8baf345c00fb7c7e72277688ce2ecc706c390498b18025428624ca25eae2e","Task_Delivery_Context.json":"cb2725244476dac2ec672c7c60585c408f664ffc5937dc19b47e1bff712d8311"}
GOLDEN={"Internal_Tooling_Training.pptx":"f2aa5b41e75a8b1a548d481be4aa27e43992d636a2fa438e4d2355a8a57ceed7","Delivery_Record.json":"8b00b6777153089a8698abf85c2b5f996f0c52d9015d7572722c3fbacb35d77c"}
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def main():
    parser=argparse.ArgumentParser(); parser.add_argument("--input-dir",type=Path,required=True); parser.add_argument("--output-dir",type=Path,required=True); parser.add_argument("--golden-dir",type=Path,required=True); args=parser.parse_args()
    for name,digest in HASHES.items():
        path=args.input_dir/name
        if not path.is_file() or sha(path)!=digest: raise SystemExit(f"冻结输入错误：{name}")
    args.output_dir.mkdir(parents=True,exist_ok=True)
    for name,digest in GOLDEN.items():
        source=args.golden_dir/name
        if not source.is_file() or sha(source)!=digest: raise SystemExit(f"Golden 错误：{name}")
        shutil.copy2(source,args.output_dir/name)
if __name__=="__main__": main()
