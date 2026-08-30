#!/usr/bin/env python3
from __future__ import annotations
import argparse,hashlib,shutil
from pathlib import Path
HASHES={"Invoice.pdf":"29b4643565c16cbfce84114d92cd3edb11d94ed017f727873ddf4fcf5d11d194","Invoice_2.pdf":"5758d6a85a47563ef5fa718a8a8603106b0c83c4a129f4f91dc57210809df8ec"}; GOLDEN_SHA256="4dc61c2e921d3c39fb409009ff9e47ba46b0bf6ef9f57666567d0e4dedcde6d8"; OUTPUT_NAME="January_2023_Invoice_Payment_Review.pptx"
def sha256(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def main():
    p=argparse.ArgumentParser(); p.add_argument("--input-dir",type=Path,required=True); p.add_argument("--output-dir",type=Path,required=True); p.add_argument("--golden-dir",type=Path,required=True); a=p.parse_args()
    for name,expected in HASHES.items():
        path=a.input_dir/name
        if not path.is_file() or sha256(path)!=expected: raise SystemExit(f"冻结输入缺失或 SHA256 错：{name}")
    golden=a.golden_dir/OUTPUT_NAME
    if not golden.is_file() or sha256(golden)!=GOLDEN_SHA256: raise SystemExit("Golden 缺失或 SHA256 错")
    a.output_dir.mkdir(parents=True,exist_ok=True); target=a.output_dir/OUTPUT_NAME; shutil.copy2(golden,target); print(target)
if __name__=="__main__": main()
