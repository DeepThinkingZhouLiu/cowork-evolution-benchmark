#!/usr/bin/env python3
"""officeval_029 的独立 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document

import upstream_verifier


TASK_ID = "cowork-evo/officeval-029-image-to-editable-training-form"
OUTPUT_NAME = "培训记录表_按图制作_可编辑.docx"
RUBRIC_VERSION = "omegause-officeval-029-adaptation-v0.1"
JUDGE_VERSION = "officeval-029-adaptation-v0.1"
LABELS = ["员工姓名", "所属部门", "性别", "任职岗位", "培训日期", "培训目标", "培训项目", "学习时长", "考核结果", "能力评定", "评定方式", "学习完成", "直属主管", "人力资源"]
SPECS = {
    "R001": ("指定 DOCX 可打开，且正文只有一张原生表格", 3, "hurdle"),
    "R002": ("表格为 8 行，区块顺序和全部标签正确", 5, "hurdle"),
    "R003": ("第 1 至 4 行的逻辑格数、合并关系和指定宽度正确", 6, "positive"),
    "R004": ("第 5 至 8 行左侧标签与右侧合并填写区正确", 4, "positive"),
    "R005": ("全部可见文字为仿宋小三", 3, "positive"),
    "R006": ("内外边框为连续黑色细实线且无斜线", 3, "positive"),
    "R007": ("交付物没有图片、截图或 OLE 压平对象", 4, "hurdle"),
    "R008": ("LibreOffice 实际渲染为单页且关键标签可提取", 3, "positive"),
    "R009": ("若整表被图片替代、缺少两行以上或缺少四个以上标签，扣 8 分", -8, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {"criterion_id": identifier, "description": description, "weight": weight, "type": kind, "passed": bool(passed), "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def render_check(path: Path) -> tuple[bool, str]:
    if not shutil.which("libreoffice") or not shutil.which("pdfinfo") or not shutil.which("pdftotext"):
        return False, "缺少渲染依赖"
    with tempfile.TemporaryDirectory(prefix="officeval029-render-") as tmp:
        proc = subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(path)], capture_output=True, text=True, timeout=180)
        pdf = Path(tmp) / f"{path.stem}.pdf"
        if proc.returncode != 0 or not pdf.is_file():
            return False, f"渲染失败：{proc.stderr[-200:]}"
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=30).stdout
        txt = Path(tmp) / "rendered.txt"
        subprocess.run(["pdftotext", str(pdf), str(txt)], capture_output=True, timeout=30)
        content = txt.read_text(encoding="utf-8", errors="ignore") if txt.is_file() else ""
        pages_one = "Pages:           1" in info or "Pages:          1" in info or "Pages: 1" in info
        ok = pages_one and all(label in content.replace("\n", "") for label in ["员工姓名", "培训日期", "人力资源意见"])
        return ok, f"单页={pages_one}，关键标签可提取={ok}"


def evaluate(artifact_dir: Path) -> dict:
    output = artifact_dir / OUTPUT_NAME
    results = []
    document = None
    try:
        document = Document(output)
    except Exception as exc:
        results.append(criterion("R001", False, f"缺失或无法打开：{exc}"))
    if document is not None:
        results.append(criterion("R001", len(document.tables) == 1, f"表格数={len(document.tables)}"))
        table = document.tables[0] if document.tables else None
        all_text = "\n".join(cell.text for row in table.rows for cell in row.cells) if table else ""
        missing = [label for label in LABELS if label not in all_text.replace("\n", "")]
        rows_ok = table is not None and len(table.rows) == 8
        results.append(criterion("R002", rows_ok and not missing, f"行数={len(table.rows) if table else 0}，缺失标签={missing}"))

        upstream = upstream_verifier.evaluate(str(artifact_dir))
        items = upstream.get("dim2_items", [])
        hits = [bool(item.get("hit")) for item in items]
        geometry_hits = hits[1:5] if len(hits) >= 5 else []
        lower_hits = hits[5:9] if len(hits) >= 9 else []
        # 上游第 1 行实现错误地要求“可见单元格少于 8”，与逐个标签/空白区交替要求矛盾；
        # 本 Judge 对第 1 行采用 8 个真实可填写单元格的忠实解释，其余几何沿用上游检测。
        row1_cells = len(table.rows[0].cells) if table else 0
        row1_ok = row1_cells == 8
        other_geometry_ok = len(geometry_hits) == 4 and all(geometry_hits[1:])
        results.append(criterion("R003", row1_ok and other_geometry_ok, f"第1行可见格={row1_cells}，上游几何命中={geometry_hits}"))
        results.append(criterion("R004", len(lower_hits) == 4 and all(lower_hits), f"第5至8行上游命中={lower_hits}"))

        font_ok = len(hits) > 10 and hits[9] and hits[10]
        border_ok = len(hits) > 11 and hits[11]
        results.append(criterion("R005", font_ok, f"字体/字号上游命中={hits[9:11] if len(hits) >= 11 else hits}"))
        results.append(criterion("R006", border_ok, f"边框上游命中={border_ok}"))

        with zipfile.ZipFile(output) as archive:
            media = [name for name in archive.namelist() if name.startswith("word/media/")]
            xml = archive.read("word/document.xml")
        flattened = bool(media) or b"<w:drawing" in xml or b"<w:pict" in xml or b"OLEObject" in xml
        results.append(criterion("R007", not flattened, f"媒体文件={len(media)}，正文图片/OLE={flattened}"))
        render_pass, render_evidence = render_check(output)
        results.append(criterion("R008", render_pass, render_evidence))
        severe = flattened or not rows_ok or len(missing) >= 4
        results.append(criterion("R009", severe, f"严重错误={severe}，上游总分={upstream.get('total_score')}/{upstream.get('max_score')}"))
    return build_result(results, output)


def build_result(results: list[dict], output: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    earned += sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, earned / total))
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
