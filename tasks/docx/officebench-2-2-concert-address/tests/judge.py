
from __future__ import annotations
import argparse, hashlib, json, re, zipfile
from collections import Counter
from pathlib import Path
from docx import Document
from openpyxl import load_workbook

WEIGHTS = {"R001": .20, "R002": .20, "R003": .16, "R004": .14, "R005": .15, "R006": .10, "R007": .05}
def c(cid, ok, evidence): return {"criterion_id": cid, "score": 1.0 if ok else 0.0, "evidence": evidence, "method": "deterministic"}
def valid_doc(path): return path.is_file() and zipfile.is_zipfile(path)
def all_text(doc): return "\n".join(p.text for p in doc.paragraphs)
def check(source_dir, out, doc, text):
    normalized = re.sub(r"\s+", " ", text).strip()
    return [
        c("R002", "John Smith" in normalized and "123 Main Street, Cityville" in normalized, "表演者姓名和地址均与 PDF 一致"),
        c("R003", "Artist" in normalized and "Address" in normalized, "姓名和地址字段标签清楚"),
        c("R004", len([p for p in doc.paragraphs if p.text.strip()]) >= 3, "标题及两个字段均为可编辑段落"),
        c("R005", not re.search(r"TODO|placeholder|Page\s+\d+", normalized, re.I), "无占位符或 PDF 页码噪声"),
    ]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'contact.docx'
    results=[]
    try:
        ok = valid_doc(out); doc = Document(out) if ok else None; text = all_text(doc) if doc else ""
        results.append(c("R001", ok and (bool(text.strip()) or bool(doc and doc.tables)), "可解析的原生 OOXML DOCX"))
        results.extend(check(source_dir, out, doc, text))
        digest = hashlib.sha256((source_dir / 'concert_post.pdf').read_bytes()).hexdigest()
        results.append(c("R006", digest == 'ae126a54d332e92eaa3c486f4aca2389e1e69db2bf89bc00935dc722e1f25932', f"输入 SHA256={digest}"))
        results.append(c("R007", bool(doc and (len(doc.paragraphs) >= 1 or doc.tables)), "输出包含可编辑段落或表格"))
    except Exception as exc:
        results = [c("R001", False, f"读取失败: {exc}")]
    vals={x["criterion_id"]: x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS)
    hurdle_failures = [key for key in ("R001", "R002") if vals.get(key, 0) != 1]
    if "R001" in hurdle_failures:
        reward = 0.0
    elif hurdle_failures:
        reward = min(reward, 0.65)
    passed=reward >= .8 and not hurdle_failures
    payload={"task_id": 'cowork-evo/officebench-2-2-concert-address', "rubric_version": 'officebench-2-2-concert-address'+"-v0.2", "judge_version": 'cowork-evo/officebench-2-2-concert-address'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed, "hurdle_failures": hurdle_failures}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
