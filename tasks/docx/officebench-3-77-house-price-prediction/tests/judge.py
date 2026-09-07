
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
    match = re.search(r"2030\D{0,80}(\d{3}(?:\.\d+)?)", compact)
    prediction = float(match.group(1)) if match else None
    trend = ("2006" in compact and "2012" in compact and "2023" in compact and any(x in compact for x in ("decline", "decrease", "fell")) and any(x in compact for x in ("rebound", "recover", "rise")))
    assumptions = any(x in compact for x in ("assumption", "assume")) and sum(x in compact for x in ("interest", "recession", "supply", "policy", "structural break")) >= 2
    return [c("R002", prediction is not None and 450 <= prediction <= 530, f"2030 预测值={prediction}，与基于 2013-2023 的线性外推一致"), c("R003", trend, "识别长期上升、2006 高点、2012 前后低谷和后续反弹"), c("R004", assumptions, "明确列出外推假设与结构性风险"), c("R005", len([p for p in doc.paragraphs if p.text.strip()]) >= 4 and "TODO" not in text, "预测、趋势和假设均为可编辑段落")]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'prediction.docx'
    results=[]
    try:
        ok = valid_doc(out); doc = Document(out) if ok else None; text = all_text(doc) if doc else ""
        results.append(c("R001", ok and (bool(text.strip()) or bool(doc and doc.tables)), "可解析的原生 OOXML DOCX"))
        results.extend(check(source_dir, out, doc, text))
        digest = hashlib.sha256((source_dir / 'house_price_index.jpg').read_bytes()).hexdigest()
        results.append(c("R006", digest == '1659f09864856d489d7b5a7ce72150787ee3e43510589f90174f5c3292cf5d14', f"输入 SHA256={digest}"))
        results.append(c("R007", bool(doc and (len(doc.paragraphs) >= 1 or doc.tables)), "输出包含可编辑段落或表格"))
    except Exception as exc:
        results = [c("R001", False, f"读取失败: {exc}")]
    vals={x["criterion_id"]: x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS)
    passed=reward >= .8 and vals.get("R001",0) == 1 and vals.get("R002",0) == 1
    payload={"task_id": 'cowork-evo/officebench-3-77-house-price-prediction', "rubric_version": 'officebench-3-77-house-price-prediction'+"-v0.2", "judge_version": 'cowork-evo/officebench-3-77-house-price-prediction'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
