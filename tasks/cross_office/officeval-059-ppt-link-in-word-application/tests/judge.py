#!/usr/bin/env python3
"""officeval_059 的 Word 单元格定位、相对链接和随附 PPT 独立 Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

from lxml import etree
from pptx import Presentation


DOC_NAME = "青梧谷灵MR共创计划_项目申报书_已链接PPT.docx"
PPT_NAME = "青梧谷灵MR共创计划.pptx"
TASK_ID = "cowork-evo/officeval-059-ppt-link-in-word-application"
RUBRIC_VERSION = "omegause-officeval-059-adaptation-v0.1"
JUDGE_VERSION = "officeval-059-adaptation-v0.1"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "r": R_NS, "rel": REL_NS}
SPECS = {
    "R001": ("两个指定交付文件均存在且可打开", 5, "hurdle"),
    "R002": ("随附 PPT 与真实输入逐字节一致，含 18 页正确项目内容", 7, "hurdle"),
    "R003": ("超链接严格位于‘项目介绍材料’右侧单元格，显示名称可识别", 7, "hurdle"),
    "R004": ("链接目标为同目录 PPT 的安全相对路径，并声明 External", 7, "hurdle"),
    "R005": ("链接为可编辑、带下划线的小五号真实 Word 超链接，而非截图", 4, "positive"),
    "R006": ("原 Word 除 document.xml 与关系表外全部包成员逐字节不变，原文字完整", 7, "hurdle"),
    "R007": ("LibreOffice 仍渲染为 4 页，第 4 页标签、链接、页眉和页码可读", 5, "positive"),
    "R008": ("申报表保持 25 行可编辑结构，最后一行边框完整且文档无批注", 4, "positive"),
    "R009": ("若链接不可点击、PPT 缺失或原表格结构破坏，扣 18 分", -18, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {
        "criterion_id": identifier,
        "score": 1.0 if passed else 0.0,
        "evidence": evidence,
        "method": "deterministic",
        "raw": {"description": description, "weight": weight, "type": kind, "passed": bool(passed)},
    }


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def locate_inputs() -> Path:
    for candidate in [Path("/app/input_files"), Path(__file__).resolve().parents[1] / "data/input_files"]:
        if (candidate / PPT_NAME).is_file() and (candidate / "青梧谷灵MR共创计划_项目申报书.docx").is_file():
            return candidate
    raise FileNotFoundError("找不到任务输入目录")


def zip_payloads(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f"ZIP 部件损坏：{bad}")
        return {name: archive.read(name) for name in archive.namelist()}


def parse_docx(path: Path) -> tuple[dict[str, bytes], etree._Element, etree._Element]:
    payloads = zip_payloads(path)
    document = etree.fromstring(payloads["word/document.xml"])
    relationships = etree.fromstring(payloads["word/_rels/document.xml.rels"])
    return payloads, document, relationships


def target_row(document) -> tuple[etree._Element | None, list]:
    for row in document.xpath(".//w:tr", namespaces=NS):
        cells = row.xpath("./w:tc", namespaces=NS)
        if len(cells) >= 2:
            label = "".join(cells[0].xpath(".//w:t/text()", namespaces=NS)).replace(" ", "")
            if "项目介绍材料" in label:
                return row, cells
    return None, []


def render_docx(path: Path) -> tuple[bool, str]:
    with tempfile.TemporaryDirectory(prefix="officeval059-judge-") as temp_dir:
        completed = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=120,
        )
        pdf = Path(temp_dir) / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf.is_file():
            return False, f"转 PDF 失败：{completed.stdout[-300:]}"
        info = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, text=True, timeout=30).stdout
        match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        text_path = Path(temp_dir) / "text.txt"
        subprocess.run(["pdftotext", "-f", "4", "-l", "4", str(pdf), str(text_path)], check=True, timeout=30)
        text = re.sub(r"\s+", "", text_path.read_text(errors="replace"))
        required = ["项目介绍材料", "青梧谷灵MR共创计划.pptx", "项目申报书", "第4页"]
        ok = pages == 4 and all(re.sub(r"\s+", "", item) in text for item in required)
        return ok, f"页数={pages}，第4页标签/链接/页眉页码可读={ok}"


def evaluate(artifact_dir: Path) -> dict:
    criteria = []
    docx = artifact_dir / DOC_NAME
    pptx = artifact_dir / PPT_NAME
    expected_names = {DOC_NAME, PPT_NAME}
    actual_names = {path.name for path in artifact_dir.iterdir() if path.is_file()} if artifact_dir.is_dir() else set()
    try:
        inputs = locate_inputs()
        output_payloads, document, relationships = parse_docx(docx)
        presentation = Presentation(pptx)
        open_ok = actual_names == expected_names
    except Exception as exc:
        for identifier in SPECS:
            criteria.append(criterion(identifier, False, f"打开失败：{exc}"))
        return make_result([path for path in (docx, pptx) if path.is_file()], criteria)

    criteria.append(criterion("R001", open_ok, f"实际交付={sorted(actual_names)}"))

    source_ppt = inputs / PPT_NAME
    ppt_same = sha256_file(pptx) == sha256_file(source_ppt)
    first_slide_text = "".join(shape.text for shape in presentation.slides[0].shapes if getattr(shape, "has_text_frame", False))
    content_ok = len(presentation.slides) == 18 and "青梧谷灵" in first_slide_text and "MR共创计划" in first_slide_text and "重构区域文化体验" in first_slide_text
    criteria.append(criterion("R002", ppt_same and content_ok, f"PPT逐字节一致={ppt_same}，页数={len(presentation.slides)}，首页内容正确={content_ok}"))

    row, cells = target_row(document)
    hyperlinks = cells[1].xpath(".//w:hyperlink", namespaces=NS) if len(cells) >= 2 else []
    display_text = "".join(hyperlinks[0].xpath(".//w:t/text()", namespaces=NS)) if len(hyperlinks) == 1 else ""
    position_ok = row is not None and len(cells) == 2 and len(hyperlinks) == 1 and "项目介绍材料PPT" in display_text and "青梧谷灵MR共创计划" in display_text
    criteria.append(criterion("R003", position_ok, f"目标行存在={row is not None}，右单元格超链接数={len(hyperlinks)}，显示={display_text!r}"))

    relation = None
    relation_id = hyperlinks[0].get(f"{{{R_NS}}}id") if hyperlinks else None
    if relation_id:
        found = relationships.xpath("./rel:Relationship[@Id=$rid]", namespaces=NS, rid=relation_id)
        relation = found[0] if len(found) == 1 else None
    target = relation.get("Target", "") if relation is not None else ""
    target_mode = relation.get("TargetMode", "") if relation is not None else ""
    pure = PurePosixPath(target) if target else None
    relative_ok = bool(target) and target == PPT_NAME and target_mode == "External" and not pure.is_absolute() and ".." not in pure.parts and ":" not in target
    criteria.append(criterion("R004", relative_ok, f"Target={target!r}，TargetMode={target_mode!r}"))

    right_cell = cells[1] if len(cells) >= 2 else None
    drawings = right_cell.xpath(".//w:drawing|.//w:object|.//w:pict", namespaces=NS) if right_cell is not None else []
    run_properties = hyperlinks[0].xpath(".//w:rPr", namespaces=NS) if hyperlinks else []
    underline = bool(run_properties and run_properties[0].find(f"{{{W_NS}}}u") is not None)
    sizes = run_properties[0].xpath("./w:sz/@w:val", namespaces=NS) if run_properties else []
    editable_ok = len(hyperlinks) == 1 and not drawings and underline and sizes == ["18"]
    criteria.append(criterion("R005", editable_ok, f"真实超链接={len(hyperlinks)==1}，静态图/对象={len(drawings)}，下划线={underline}，字号半点={sizes}"))

    source_doc = inputs / "青梧谷灵MR共创计划_项目申报书.docx"
    source_payloads, source_document, _ = parse_docx(source_doc)
    changed_members = sorted(
        name for name in set(source_payloads) | set(output_payloads)
        if source_payloads.get(name) != output_payloads.get(name)
    )
    source_text = "".join(source_document.xpath(".//w:t/text()", namespaces=NS))
    output_text = "".join(document.xpath(".//w:t/text()", namespaces=NS))
    preservation_ok = changed_members == ["word/_rels/document.xml.rels", "word/document.xml"] and output_text == source_text + display_text
    criteria.append(criterion("R006", preservation_ok, f"变化部件={changed_members}，原文字完整={output_text == source_text + display_text}"))

    render_ok, render_evidence = render_docx(docx)
    criteria.append(criterion("R007", render_ok, render_evidence))

    rows = document.xpath(".//w:tbl[2]/w:tr", namespaces=NS)
    last_borders = cells[0].xpath("./w:tcPr/w:tcBorders/w:bottom", namespaces=NS) + cells[1].xpath("./w:tcPr/w:tcBorders/w:bottom", namespaces=NS) if len(cells) >= 2 else []
    comments = [name for name in output_payloads if name.startswith("word/comments")]
    table_ok = len(rows) == 25 and len(last_borders) == 2 and not comments
    criteria.append(criterion("R008", table_ok, f"申报表行数={len(rows)}，末行底边={len(last_borders)}，批注部件={comments}"))

    severe = not relative_ok or not ppt_same or not table_ok
    criteria.append(criterion("R009", not severe, f"严重错误={severe}"))
    return make_result([docx, pptx], criteria)


def make_result(paths: list[Path], criteria: list[dict]) -> dict:
    hurdles = [item["criterion_id"] for item in criteria if item["raw"]["type"] == "hurdle" and not item["raw"]["passed"]]
    total = sum(item["raw"]["weight"] for item in criteria if item["raw"]["weight"] > 0)
    earned = sum(item["raw"]["weight"] for item in criteria if item["raw"]["weight"] > 0 and item["raw"]["passed"])
    penalties = sum(abs(item["raw"]["weight"]) for item in criteria if item["raw"]["weight"] < 0 and not item["raw"]["passed"])
    reward = 0.0 if hurdles else max(0.0, min(1.0, (earned - penalties) / total))
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [sha256_file(path) for path in paths if path.is_file()],
        "criterion_results": criteria,
        "reward": reward,
        "passed": reward >= .8 and not hurdles,
        "hurdle_failures": hurdles,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    evaluation = evaluate(args.artifact_dir)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(evaluation, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.write_text(f"{evaluation['reward']:.6f}\n", encoding="utf-8")
    print(json.dumps(evaluation, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
