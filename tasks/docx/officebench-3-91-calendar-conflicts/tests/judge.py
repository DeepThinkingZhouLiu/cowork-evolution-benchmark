
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
    event_pairs = [
        ("bob sleeping", "tom class"), ("bob lunch", "tom lunch with david"),
        ("bob nap", "tom lunch with david"), ("bob class", "tom class"),
        ("bob dinner", "tom dinner"),
    ]
    ranges = ("08:00-11:00", "12:30-13:00", "13:00-13:30", "16:00-18:00", "18:30-19:30")
    pair_hits = sum(a in compact and b in compact for a, b in event_pairs)
    range_hits = sum(x in compact for x in ranges)
    count_signal = bool(re.search(r"(?:5|five)\s+(?:conflict|overlap)", compact)) or pair_hits == 5
    return [c("R002", pair_hits == 5 and range_hits == 5, f"五组冲突事件与交叉时段均准确：事件={pair_hits}/5，时段={range_hits}/5"), c("R003", count_signal, "冲突数量明确为 5"), c("R004", "14:00" in compact and any(x in compact for x in ("not counted", "not a conflict", "not overlap")), "未把 14:00 端点相接误判为重叠"), c("R005", len([p for p in doc.paragraphs if p.text.strip()]) >= 6 and "TODO" not in text, "每组冲突为可编辑条目")]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'schedule.docx'
    results=[]
    try:
        ok = valid_doc(out); doc = Document(out) if ok else None; text = all_text(doc) if doc else ""
        results.append(c("R001", ok and (bool(text.strip()) or bool(doc and doc.tables)), "可解析的原生 OOXML DOCX"))
        results.extend(check(source_dir, out, doc, text))
        digests = {name: hashlib.sha256((source_dir / name).read_bytes()).hexdigest() for name in ('Bob.ics', 'Tom.ics')}
        expected = {'Bob.ics': 'a641a03140fc0903fb3d3a320642c146edc5689ffae06712c84cb4b5ee60be98', 'Tom.ics': '4f946c13f4ee53deb6d74b70f73da42cc9a1063fc1183b7daca4a44bf69eba99'}
        results.append(c("R006", digests == expected, f"两份日历快照匹配={digests == expected}"))
        results.append(c("R007", bool(doc and (len(doc.paragraphs) >= 1 or doc.tables)), "输出包含可编辑段落或表格"))
    except Exception as exc:
        results = [c("R001", False, f"读取失败: {exc}")]
    vals={x["criterion_id"]: x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS)
    passed=reward >= .8 and vals.get("R001",0) == 1 and vals.get("R002",0) == 1
    payload={"task_id": 'cowork-evo/officebench-3-91-calendar-conflicts', "rubric_version": 'officebench-3-91-calendar-conflicts'+"-v0.2", "judge_version": 'cowork-evo/officebench-3-91-calendar-conflicts'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
