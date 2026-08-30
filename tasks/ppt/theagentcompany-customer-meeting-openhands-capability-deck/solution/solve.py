#!/usr/bin/env python3
from __future__ import annotations
import argparse,hashlib,shutil
from pathlib import Path

HASHES={"Original_Request.md":"eb8a006788c291de4c76a47e8f72a33670395d1f2999c99b45c7e0dd8c5bbdef","RocketChat_Recipient_Context.json":"865c354cfc784ef51994a0b0159dd396f3a76c3f8ba01046a039d3d0d1ac07b8","OpenHands_Evaluation_Snapshot.json":"aa411f23fbb791720b0c9274a5900685c13d929f2dce1c20d711a406a26d3338"}
GOLDEN={"openhands_intro.pptx":"171ed28bbc45f0086240c61f9dafc4031e8f5063ebb77058264982988258245e","Delivery_Record.json":"61ef42069eff2a959d9a74afdd985c7d4fd314d304e4ad38ae97e0056497d439"}
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
