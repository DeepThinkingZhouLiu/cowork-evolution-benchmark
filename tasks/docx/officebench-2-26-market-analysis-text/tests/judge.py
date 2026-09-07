
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

    import subprocess
    src = next(source_dir.glob("*.pdf")); expected = subprocess.check_output(["pdftotext","-layout",str(src),"-"], text=True, errors="replace").strip()
    expected_norm = re.sub(r"\s+", " ", expected).strip(); actual_norm = re.sub(r"\s+", " ", text).strip()
    anchors = ("Market Analysis and Marketing Plan", "Who is Your Customer?", "How Large is Your Potential Market?", "What Would they be Willing to Spend")
    return [c("R002", actual_norm == expected_norm, "PDF 全文归一化后逐字匹配"), c("R003", all(x in actual_norm for x in anchors), "四个市场分析主题齐全"), c("R004", len(actual_norm) == len(expected_norm), "未截断或添加无关文本"), c("R005", len(doc.paragraphs) >= 1 and "placeholder" not in text.casefold(), "文本为可编辑段落且无占位符")]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'market_analysis.docx'
    results=[]
    try:
        ok = valid_doc(out); doc = Document(out) if ok else None; text = all_text(doc) if doc else ""
        results.append(c("R001", ok and (bool(text.strip()) or bool(doc and doc.tables)), "可解析的原生 OOXML DOCX"))
        results.extend(check(source_dir, out, doc, text))
        digest = hashlib.sha256((source_dir / 'market_analysis.pdf').read_bytes()).hexdigest()
        results.append(c("R006", digest == '8e39f62f9fb5747c5ac1527eea0162ef7f91a5f178a9890eb46a995ef5d5207c', f"输入 SHA256={digest}"))
        results.append(c("R007", bool(doc and (len(doc.paragraphs) >= 1 or doc.tables)), "输出包含可编辑段落或表格"))
    except Exception as exc:
        results = [c("R001", False, f"读取失败: {exc}")]
    vals={x["criterion_id"]: x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS)
    passed=reward >= .8 and vals.get("R001",0) == 1 and vals.get("R002",0) == 1
    payload={"task_id": 'cowork-evo/officebench-2-26-market-analysis-text', "rubric_version": 'officebench-2-26-market-analysis-text'+"-v0.2", "judge_version": 'cowork-evo/officebench-2-26-market-analysis-text'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
