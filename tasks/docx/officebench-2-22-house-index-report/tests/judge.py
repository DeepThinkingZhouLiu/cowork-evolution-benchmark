
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
    compact = re.sub(r"\s+", " ", text).casefold()
    anchors = ["los angeles county", "index 2000=100", "1975", "16.91", "2006", "255.34", "2012", "165.26", "2023", "375.2"]
    predictions = [float(value) for value in re.findall(r"2030[^\d]{0,40}(\d+(?:\.\d+)?)", compact)]
    trend_ok = all(any(word in compact for word in group) for group in [
        ("long-run increase", "long-term increase", "long-run rise", "upward trend"),
        ("decline", "fell", "decrease"),
        ("rebound", "recovery", "rose again"),
    ])
    assumption_ok = "assumption" in compact or "assuming" in compact
    return [
        c("R002", all(anchor in compact for anchor in anchors), f"源数据锚点覆盖={sum(anchor in compact for anchor in anchors)}/{len(anchors)}"),
        c("R003", trend_ok, "识别长期上升、2006–2012 回落与此后回升"),
        c("R004", any(250 <= value <= 700 for value in predictions) and assumption_ok, f"2030 预测值={predictions}，有假设={assumption_ok}"),
        c("R005", len([p for p in doc.paragraphs if p.text.strip()]) >= 4 and not re.search(r"TODO|placeholder|\bnan\b", text, re.I), "报告为多段可编辑文本且无占位内容"),
    ]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'house_price_index.docx'
    results=[]
    try:
        ok = valid_doc(out); doc = Document(out) if ok else None; text = all_text(doc) if doc else ""
        results.append(c("R001", ok and (bool(text.strip()) or bool(doc and doc.tables)), "可解析的原生 OOXML DOCX"))
        results.extend(check(source_dir, out, doc, text))
        digest = hashlib.sha256((source_dir / 'house_price_index.xlsx').read_bytes()).hexdigest()
        results.append(c("R006", digest == 'e2cc3d30e3fd218d4a8c8c321a1752c4e2ad403e040f04c87750eda76ad58adf', f"输入 SHA256={digest}"))
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
    payload={"task_id": 'cowork-evo/officebench-2-22-house-index-report', "rubric_version": 'officebench-2-22-house-index-report'+"-v0.2", "judge_version": 'cowork-evo/officebench-2-22-house-index-report'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed, "hurdle_failures": hurdle_failures}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
