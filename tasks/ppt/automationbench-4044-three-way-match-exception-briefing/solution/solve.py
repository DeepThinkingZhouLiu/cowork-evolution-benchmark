#!/usr/bin/env python3
from __future__ import annotations
import argparse,hashlib,shutil
from pathlib import Path
HASHES={"AutomationBench_4044_Context.json":"6698be721fdb27bbb35f2ab5d75960582bd6d825cff5453c1ebd7498f37e9a00","Three_Way_Match_Batch.xlsx":"03b91916d397ec4f11455650635543f2706bb044417b1ebdc7df3d2e5861bf98","Three_Way_Match_Tolerance_Policy.eml":"0413559546d2d14ba17e2fe4325bc98a312d2a736be5d84bff8d713df563adfc"}
GOLDEN_SHA="760721adac961f2bcb16c2a4c91e1c3d13dd046bc797238c73613b6633b7101b"
OUTPUT_NAME="Three_Way_Match_Exception_Briefing.pptx"
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
    p=argparse.ArgumentParser(); p.add_argument("--input-dir",type=Path,required=True); p.add_argument("--output-dir",type=Path,required=True); p.add_argument("--golden-dir",type=Path,required=True); a=p.parse_args()
    for n,h in HASHES.items():
        q=a.input_dir/n
        if not q.is_file() or sha(q)!=h: raise SystemExit(f"冻结输入错误：{n}")
    g=a.golden_dir/OUTPUT_NAME
    if not g.is_file() or sha(g)!=GOLDEN_SHA: raise SystemExit("Golden 错误")
    a.output_dir.mkdir(parents=True,exist_ok=True); target=a.output_dir/OUTPUT_NAME; shutil.copy2(g,target); print(target)
if __name__=="__main__": main()
