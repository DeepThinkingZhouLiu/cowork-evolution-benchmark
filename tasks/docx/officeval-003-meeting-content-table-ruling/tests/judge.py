#!/usr/bin/env python3
"""officeval_003 的独立 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

import upstream_verifier


TASK_ID = "cowork-evo/officeval-003-meeting-content-table-ruling"
OUTPUT_NAME = "部门协调会记录_会议内容横线表格版.docx"
RUBRIC_VERSION = "omegause-officeval-003-adaptation-v0.1"
JUDGE_VERSION = "officeval-003-adaptation-v0.1"
EXPECTED_HEADER = ["部门协调会", "2026年3月18日", "许承柏", "林若棠", "综合楼二层会议室", "陈予宁、侯景澜、魏预、周言朔", "会议内容"]
EXPECTED_FIRST = "许承柏：部门共有5人，本次到会5人，符合会议人员参会要求。经与陈予宁沟通，建议侯景澜、魏预列"
EXPECTED_LAST = "许承柏：今天会议议程已全部结束，散会！"

SPECS = {
    "R001": ("指定 DOCX 存在、可打开且只有一张会议内容表", 3, "hurdle"),
    "R002": ("表格恰含 95 条非空会议记录，没有空白行", 5, "hurdle"),
    "R003": ("每条记录下方都有贯穿整行的黑色细实线", 5, "positive"),
    "R004": ("相邻横线间距一致，最大误差不超过 0.15 厘米", 4, "positive"),
    "R005": ("标题、时间、人员、地点和出席信息完整保留", 4, "positive"),
    "R006": ("95 条非空会议内容逐字保留且顺序不变", 6, "hurdle"),
    "R007": ("文档不含图片或压平内容，会议文字仍可编辑", 3, "positive"),
    "R008": ("LibreOffice 可实际渲染，首页和末页均有可提取文本", 3, "positive"),
    "R009": ("若出现空白行、丢失关键记录或插入图片，扣 8 分", -8, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {"criterion_id": identifier, "description": description, "weight": weight, "type": kind, "passed": bool(passed), "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def table_texts(document: Document) -> list[str]:
    if len(document.tables) != 1:
        return []
    return ["".join(cell.text for cell in row.cells).strip() for row in document.tables[0].rows]


def borders_ok(document: Document) -> tuple[bool, str]:
    failures = []
    for index, row in enumerate(document.tables[0].rows, 1):
        cell = row.cells[0]
        tc_pr = cell._tc.tcPr
        borders = tc_pr.find(qn("w:tcBorders")) if tc_pr is not None else None
        bottom = borders.find(qn("w:bottom")) if borders is not None else None
        if bottom is None or bottom.get(qn("w:val")) != "single" or bottom.get(qn("w:color"), "").upper() not in {"000000", "AUTO", ""}:
            failures.append(index)
            continue
        try:
            if int(bottom.get(qn("w:sz"), "99")) > 6:
                failures.append(index)
        except ValueError:
            failures.append(index)
    return not failures, "95 行均为黑色细实线" if not failures else f"边框不合格行={failures[:8]}"


def row_spacing_ok(document: Document) -> tuple[bool, str]:
    heights = []
    for row in document.tables[0].rows:
        tr_pr = row._tr.trPr
        h = tr_pr.find(qn("w:trHeight")) if tr_pr is not None else None
        if h is None or h.get(qn("w:val")) is None:
            return False, "至少一行没有显式行高"
        heights.append(int(h.get(qn("w:val"))) / 567.0)
    spread = max(heights) - min(heights) if heights else 99.0
    return spread <= 0.15, f"行高范围={min(heights):.3f}–{max(heights):.3f}cm，误差={spread:.3f}cm"


def has_images(document: Document) -> bool:
    return bool(document.inline_shapes or document.part.element.findall(".//" + qn("w:drawing")) or document.part.element.findall(".//" + qn("w:pict")))


def render_check(path: Path) -> tuple[bool, str]:
    if shutil.which("libreoffice") is None or shutil.which("pdftotext") is None:
        return False, "缺少 LibreOffice 或 pdftotext"
    with tempfile.TemporaryDirectory(prefix="officeval003-render-") as tmp:
        proc = subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(path)], capture_output=True, text=True, timeout=180)
        pdf = Path(tmp) / f"{path.stem}.pdf"
        if proc.returncode != 0 or not pdf.is_file():
            return False, f"渲染失败：{proc.stderr[-300:]}"
        txt = Path(tmp) / "rendered.txt"
        proc = subprocess.run(["pdftotext", str(pdf), str(txt)], capture_output=True, text=True, timeout=60)
        content = txt.read_text(encoding="utf-8", errors="ignore") if txt.is_file() else ""
        ok = proc.returncode == 0 and "部门协调会" in content and "散会" in content
        return ok, f"PDF={pdf.stat().st_size}B，首页标题与末条记录可提取={ok}"


def evaluate(artifact_dir: Path) -> dict:
    output = artifact_dir / OUTPUT_NAME
    results = []
    document = None
    if output.is_file():
        try:
            document = Document(output)
        except Exception as exc:
            results.append(criterion("R001", False, f"无法打开：{exc}"))
    else:
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))

    if document is not None:
        results.append(criterion("R001", len(document.tables) == 1, f"表格数={len(document.tables)}"))
        texts = table_texts(document)
        no_blanks = len(texts) == 95 and all(texts)
        results.append(criterion("R002", no_blanks, f"行数={len(texts)}，空行数={sum(not value for value in texts)}"))

        border_pass, border_evidence = borders_ok(document) if len(document.tables) == 1 else (False, "没有唯一目标表格")
        results.append(criterion("R003", border_pass, border_evidence))
        spacing_pass, spacing_evidence = row_spacing_ok(document) if len(document.tables) == 1 else (False, "没有唯一目标表格")
        results.append(criterion("R004", spacing_pass, spacing_evidence))

        full_text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        missing_header = [value for value in EXPECTED_HEADER if value not in full_text]
        results.append(criterion("R005", not missing_header, f"缺失顶部字段={missing_header}"))

        source_order_ok = len(texts) == 95 and texts[0] == EXPECTED_FIRST and texts[-1] == EXPECTED_LAST
        try:
            upstream = upstream_verifier.evaluate(str(artifact_dir))
            upstream_details = upstream.get("dim2_items", [])
            upstream_score = upstream.get("total_score", upstream.get("dim2_score", 0))
        except Exception as exc:
            upstream_details = []
            upstream_score = f"error:{exc}"
        source_order_ok = source_order_ok and "侯景澜：上月两项重点现场事项整改推进如下：" not in full_text
        results.append(criterion("R006", source_order_ok, f"首尾与数量校验={source_order_ok}，上游分数={upstream_score}"))

        image_present = has_images(document)
        results.append(criterion("R007", not image_present, f"内联图片={len(document.inline_shapes)}，存在 drawing/pict={image_present}"))
        render_pass, render_evidence = render_check(output)
        results.append(criterion("R008", render_pass, render_evidence))
        severe = (not no_blanks) or (not source_order_ok) or image_present
        results.append(criterion("R009", severe, f"空行/内容/图片严重错误={severe}；上游明细数={len(upstream_details)}"))

    return build_result(results, output)


def build_result(results: list[dict], output: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdles = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdles:
        reward = 0.0
    normalized = [{"criterion_id": item["criterion_id"], "score": (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"], "evidence": item["evidence"], "method": item["method"], "raw": {"description": item["description"], "weight": item["weight"], "type": item["type"], "passed": item["passed"]}} for item in results]
    digest = hashlib.sha256(output.read_bytes() if output.is_file() else b"").hexdigest()
    return {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [digest], "criterion_results": normalized, "reward": round(reward, 6), "passed": reward >= 0.8, "hurdle_failures": hurdles}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
