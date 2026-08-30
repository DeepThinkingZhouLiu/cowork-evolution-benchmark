#!/usr/bin/env python3
from __future__ import annotations
import argparse,hashlib,shutil
from pathlib import Path

HASHES={"Original_Request.md":"c8f04427aa726c69a25ab593c0ef65ea6a0a9503d79c9138d6b4afa296d57a37","GitLab_Repository_Snapshot.json":"b48461179a363e8b70cb92d3692ed7dc53ae67cbfe76984627dc04c8c597a729"}
GOLDEN="ad38f21f0a2be3b6295e6911bf5e583963333b1cb661457864ece532d9790771"

def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    parser=argparse.ArgumentParser(); parser.add_argument("--input-dir",type=Path,required=True); parser.add_argument("--output-dir",type=Path,required=True); parser.add_argument("--golden-dir",type=Path,required=True); args=parser.parse_args()
    for name,digest in HASHES.items():
        path=args.input_dir/name
        if not path.is_file() or sha(path)!=digest: raise SystemExit(f"冻结输入错误：{name}")
    source=args.golden_dir/"Engineering_Progress.pptx"
    if not source.is_file() or sha(source)!=GOLDEN: raise SystemExit("Golden 错误")
    args.output_dir.mkdir(parents=True,exist_ok=True); shutil.copy2(source,args.output_dir/source.name)

if __name__=="__main__": main()
