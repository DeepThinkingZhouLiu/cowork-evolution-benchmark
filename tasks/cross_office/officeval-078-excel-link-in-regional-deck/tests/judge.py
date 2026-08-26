#!/usr/bin/env python3
"""officeval_078 的第 8 页形状级链接、几何位置和随附工作簿独立 Judge。"""

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
from openpyxl import load_workbook
from pptx import Presentation


PPT_NAME = "区域协同汇报_第8页Excel超链接版.pptx"
XLSX_NAME = "CRM_pipeline_restyled_reordered.xlsx"
TASK_ID = "cowork-evo/officeval-078-excel-link-in-regional-deck"
RUBRIC_VERSION = "omegause-officeval-078-adaptation-v0.1"
JUDGE_VERSION = "officeval-078-adaptation-v0.1"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P_NS, "a": A_NS, "r": R_NS, "rel": REL_NS}
SPECS = {
    "R001": ("两个指定交付文件均存在且可打开，PPT 保持 10 页宽屏", 5, "hurdle"),
    "R002": ("随附工作簿与真实输入逐字节一致，三张工作表及记录结构完整", 7, "hurdle"),
    "R003": ("第 8 页仅有一个形状级超链接按钮，文字明确表达打开 CRM Excel", 7, "hurdle"),
    "R004": ("按钮位于右侧空白区且尺寸合规，不遮挡原有表格、标题和页脚", 6, "positive"),
    "R005": ("链接目标为同目录工作簿的安全相对路径，并声明 External", 7, "hurdle"),
    "R006": ("原 PPT 仅第 8 页 XML 与其关系表变化，原 10 页内容和对象保持不变", 7, "hurdle"),
    "R007": ("LibreOffice 可渲染为 10 页，第 8 页标题、表格、按钮和页码可读", 5, "positive"),
    "R008": ("按钮为可编辑圆角形状，采用 Excel 语义配色且无批注或多余交付", 4, "positive"),
    "R009": ("若链接不可访问、工作簿缺失、页数变化或原页面严重破坏，扣 18 分", -18, "penalty"),
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
        if (candidate / XLSX_NAME).is_file() and (candidate / "区域协同汇报.pptx").is_file():
            return candidate
    raise FileNotFoundError("找不到任务输入目录")


def zip_payloads(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f"ZIP 部件损坏：{bad}")
        return {name: archive.read(name) for name in archive.namelist()}


def shape_signature(shape) -> tuple:
    text = shape.text if getattr(shape, "has_text_frame", False) else ""
    return (int(shape.shape_type), shape.left, shape.top, shape.width, shape.height, text)


def render(path: Path) -> tuple[bool, str]:
    with tempfile.TemporaryDirectory(prefix="officeval078-judge-") as temp_dir:
        completed = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=150,
        )
        pdf = Path(temp_dir) / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf.is_file():
            return False, f"转 PDF 失败：{completed.stdout[-300:]}"
        info = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, text=True, timeout=30).stdout
        match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        text_path = Path(temp_dir) / "text.txt"
        subprocess.run(["pdftotext", "-f", "8", "-l", "8", str(pdf), str(text_path)], check=True, timeout=30)
        text = re.sub(r"\s+", "", text_path.read_text(errors="replace"))
        required = ["StrategicPipelineStatusUpdate", "Pipeline", "Q2action", "打开Excel", "CRM工作簿", "08/10"]
        ok = pages == 10 and all(re.sub(r"\s+", "", value) in text for value in required)
        return ok, f"页数={pages}，第8页标题/表格/按钮/页码可读={ok}"


def evaluate(artifact_dir: Path) -> dict:
    criteria = []
    pptx = artifact_dir / PPT_NAME
    xlsx = artifact_dir / XLSX_NAME
    expected_names = {PPT_NAME, XLSX_NAME}
    actual_names = {path.name for path in artifact_dir.iterdir() if path.is_file()} if artifact_dir.is_dir() else set()
    try:
        inputs = locate_inputs()
        source_presentation = Presentation(inputs / "区域协同汇报.pptx")
        presentation = Presentation(pptx)
        workbook = load_workbook(xlsx, data_only=False, read_only=False)
        output_payloads = zip_payloads(pptx)
        source_payloads = zip_payloads(inputs / "区域协同汇报.pptx")
    except Exception as exc:
        for identifier in SPECS:
            criteria.append(criterion(identifier, False, f"打开失败：{exc}"))
        return make_result([path for path in (pptx, xlsx) if path.is_file()], criteria)

    page_ok = len(presentation.slides) == 10 and abs(presentation.slide_width / 360000 - 33.867) <= .05 and abs(presentation.slide_height / 360000 - 19.05) <= .05
    criteria.append(criterion("R001", actual_names == expected_names and page_ok, f"实际交付={sorted(actual_names)}，页数={len(presentation.slides)}，宽屏={page_ok}"))

    workbook_same = sha256_file(xlsx) == sha256_file(inputs / XLSX_NAME)
    dimensions = [(sheet.title, sheet.max_row, sheet.max_column) for sheet in workbook.worksheets]
    expected_dimensions = [("Strategic Pipeline", 160, 7), ("Lead Development", 220, 24), ("Client Portfolio", 480, 43)]
    workbook_ok = workbook_same and dimensions == expected_dimensions
    criteria.append(criterion("R002", workbook_ok, f"逐字节一致={workbook_same}，工作表维度={dimensions}"))

    slide_root = etree.fromstring(output_payloads["ppt/slides/slide8.xml"])
    relationships = etree.fromstring(output_payloads["ppt/slides/_rels/slide8.xml.rels"])
    links = slide_root.xpath(".//p:cNvPr/a:hlinkClick", namespaces=NS)
    link_parent = links[0].getparent() if len(links) == 1 else None
    linked_shape = link_parent.getparent().getparent() if link_parent is not None else None
    button_text = "".join(linked_shape.xpath(".//a:t/text()", namespaces=NS)) if linked_shape is not None else ""
    shape_hyperlink_ok = len(links) == 1 and "打开Excel" in button_text.replace(" ", "") and "CRM工作簿" in button_text.replace(" ", "")
    criteria.append(criterion("R003", shape_hyperlink_ok, f"形状级链接数={len(links)}，按钮文字={button_text!r}"))

    xfrm = linked_shape.find("./p:spPr/a:xfrm", namespaces=NS) if linked_shape is not None else None
    if xfrm is not None:
        offset = xfrm.find("a:off", namespaces=NS)
        extent = xfrm.find("a:ext", namespaces=NS)
        geometry = tuple(int(offset.get(key)) / 360000 for key in ("x", "y")) + tuple(int(extent.get(key)) / 360000 for key in ("cx", "cy"))
    else:
        geometry = (0, 0, 0, 0)
    x, y, width, height = geometry
    geom_ok = 26 <= x <= 29 and 5 <= y <= 11 and 3.5 <= width <= 5.5 and 1.5 <= height <= 3.5 and x >= 25.65 and x + width <= 32.17 and y + height < 14.73
    criteria.append(criterion("R004", geom_ok, f"按钮几何={tuple(round(value, 2) for value in geometry)}cm，表格右缘=25.65cm"))

    relation_id = links[0].get(f"{{{R_NS}}}id") if links else None
    relation_matches = relationships.xpath("./rel:Relationship[@Id=$rid]", namespaces=NS, rid=relation_id) if relation_id else []
    relation = relation_matches[0] if len(relation_matches) == 1 else None
    target = relation.get("Target", "") if relation is not None else ""
    mode = relation.get("TargetMode", "") if relation is not None else ""
    relation_type = relation.get("Type", "") if relation is not None else ""
    pure = PurePosixPath(target) if target else None
    relative_ok = target == XLSX_NAME and mode == "External" and relation_type.endswith("/hyperlink") and not pure.is_absolute() and ".." not in pure.parts and ":" not in target and (artifact_dir / target).is_file()
    criteria.append(criterion("R005", relative_ok, f"Target={target!r}，TargetMode={mode!r}，关系类型={relation_type.rsplit('/', 1)[-1]!r}"))

    changed_members = sorted(
        name for name in set(source_payloads) | set(output_payloads)
        if source_payloads.get(name) != output_payloads.get(name)
    )
    original_objects_ok = True
    for source_slide, output_slide in zip(source_presentation.slides, presentation.slides):
        expected = [shape_signature(shape) for shape in source_slide.shapes]
        actual = [shape_signature(shape) for shape in list(output_slide.shapes)[:len(expected)]]
        if expected != actual:
            original_objects_ok = False
            break
    preservation_ok = changed_members == ["ppt/slides/_rels/slide8.xml.rels", "ppt/slides/slide8.xml"] and original_objects_ok
    criteria.append(criterion("R006", preservation_ok, f"变化部件={changed_members}，原对象保持={original_objects_ok}"))

    render_ok, render_evidence = render(pptx)
    criteria.append(criterion("R007", render_ok, render_evidence))

    geometry_type = linked_shape.xpath("./p:spPr/a:prstGeom/@prst", namespaces=NS) if linked_shape is not None else []
    fills = linked_shape.xpath("./p:spPr/a:solidFill/a:srgbClr/@val", namespaces=NS) if linked_shape is not None else []
    comments = [name for name in output_payloads if "comment" in name.lower()]
    style_ok = geometry_type == ["roundRect"] and fills == ["107C41"] and not comments and actual_names == expected_names
    criteria.append(criterion("R008", style_ok, f"形状={geometry_type}，填充={fills}，批注部件={comments}"))

    severe = not relative_ok or not workbook_same or not page_ok or not original_objects_ok
    criteria.append(criterion("R009", not severe, f"严重错误={severe}"))
    return make_result([pptx, xlsx], criteria)


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
