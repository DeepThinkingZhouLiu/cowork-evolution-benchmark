#!/usr/bin/env python3
from __future__ import annotations
import argparse,hashlib,shutil
from pathlib import Path
HASHES={"Original_Request.md":"621281d20fa3e0b4dbf6903ff871e035c99bb194a80b4fe3e1188f7f7954281b","stock_data_raw.csv":"5e83e1084cdebcfe78892cc345d38a8d584d2533e58efab722294a94edf1874c"}
GOLDEN={"stock_cleaned.csv":"1b8940d298205bed6700b1f7aea0321385c5aff0599f3fcec173300c9bb3fcbf","stock_analysis.pptx":"3994c9b535ec19fb779a147c56804a27c3ed5837ae1ff839dc87847e9ae4ed5a","stock_analysis_script.md":"58749f5bb5bddc1504b5585a47f6772dfd734a879e3aa3027733cc65ca718a07"}
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
