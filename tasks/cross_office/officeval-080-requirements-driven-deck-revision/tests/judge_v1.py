#!/usr/bin/env python3
"""officeval_080 离线 Harbor Judge：从两份真实输入独立重算真值。"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import os
import re
import subprocess
import sys
import tempfile
from dataclasses import asdict, dataclass
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from lxml import etree
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

sys.dont_write_bytecode = True


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}
EMU_PER_CM = 360000.0
OUTPUT_PPTX = "公开课课件_修改完成.pptx"
OUTPUT_XLSX = "修订响应.xlsx"
OUTPUT_PROCESSOR = "revision_requirements_processor.py"
OUTPUT_PLAN = "revision_plan.json"
REQUIRED_FILES = [OUTPUT_PPTX, OUTPUT_XLSX, OUTPUT_PROCESSOR, OUTPUT_PLAN]
APPROVED_EFFECTS = {"fade", "wipe(left)", "fly(inFromRight)", "checkerboard(across)"}
EXPECTED_ACTIONS = {
    "R01": "increase_cover_fonts",
    "R02": "match_background",
    "R03": "insert_review_image_right",
    "R04": "image_left_text_right",
    "R05": "insert_two_review_images",
    "R06": "replace_arrow_answers_with_two_lines",
    "R07": "animate_heading_and_content_with_distinct_effects",
    "R08": "animate_heading_and_content_with_distinct_effects",
    "R09": "replace_slide_with_review_image",
    "R10": "animate_all_non_title_text_same_effect",
    "R11": "remove_bottom_right_image",
    "R12": "unique_transition_each_slide",
}
EXPECTED_RESPONSE_CHECKS = {
    "R01": "pptx:slide=1;check=font_delta_2pt",
    "R02": "pptx:slide=2;check=background_match_slide3",
    "R03": "pptx:slide=3;check=embedded_image_size_position",
    "R04": "pptx:slide=4;check=image_left_text_right",
    "R05": "pptx:slide=5;check=two_embedded_images",
    "R06": "pptx:slide=6;check=two_lines_after_each_arrow",
    "R07": "pptx:slide=8;check=distinct_group_animations",
    "R08": "pptx:slide=9;check=distinct_group_animations",
    "R09": "pptx:slide=10;check=replaced_title_and_curve",
    "R10": "pptx:slide=11;check=same_click_animations",
    "R11": "pptx:slide=16;check=bottom_right_image_removed",
    "R12": "pptx:slides=1-18;check=unique_transitions",
}
CN_DIGITS = {
    "零": 0,
    "一": 1,
    "二": 2,
    "两": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
}


@dataclass
class Check:
    id: str
    description: str
    passed: bool
    detail: str
    weight: float = 1.0
    hurdle: bool = False


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def chinese_number(value: str) -> int:
    value = value.strip()
    if value.isdigit():
        return int(value)
    if value == "十":
        return 10
    if "十" in value:
        left, right = value.split("十", 1)
        return CN_DIGITS.get(left, 1) * 10 + CN_DIGITS.get(right, 0)
    return CN_DIGITS[value]


def page_mentions(text: str) -> list[int]:
    pages: list[int] = []
    for raw in re.findall(r"第([零一二两三四五六七八九十\d]+)页", text):
        page = chinese_number(raw)
        if page not in pages:
            pages.append(page)
    return pages


def first_index(paragraphs: list[str], pattern: str) -> int:
    compiled = re.compile(pattern)
    for index, text in enumerate(paragraphs):
        if compiled.search(text):
            return index
    raise ValueError(f"输入要求缺少模式：{pattern}")


def input_truth(requirements_path: Path, source_pptx_path: Path) -> dict:
    document = Document(requirements_path)
    paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs]
    assets: list[dict] = []
    for paragraph_index, paragraph in enumerate(document.paragraphs):
        for drawing in paragraph._p.xpath(".//w:drawing"):
            relationship_ids = drawing.xpath(".//a:blip/@r:embed")
            extents = drawing.xpath(".//wp:extent")
            if len(relationship_ids) != 1 or len(extents) != 1:
                continue
            part = document.part.related_parts[relationship_ids[0]]
            width_cm = int(extents[0].get("cx")) / EMU_PER_CM
            height_cm = int(extents[0].get("cy")) / EMU_PER_CM
            with Image.open(BytesIO(part.blob)) as image:
                pixels = image.size
            assets.append(
                {
                    "order": len(assets) + 1,
                    "paragraph": paragraph_index,
                    "sha256": sha256_bytes(part.blob),
                    "size_bytes": len(part.blob),
                    "display_ratio": width_cm / height_cm,
                    "pixels": pixels,
                }
            )

    cover = first_index(paragraphs, r"第一页是别人做的.*第二页")
    slide3 = first_index(paragraphs, r"第三页.*图宽")
    slide4 = first_index(paragraphs, r"第四页.*图片放左侧")
    slide5_title = first_index(paragraphs, r"第五页.*What is the story")
    slide5_action = first_index(paragraphs, r"加上两幅图")
    slide5_asset = next(asset["paragraph"] for asset in assets if asset["paragraph"] > slide5_action)
    slide6 = first_index(paragraphs, r"第六页.*箭头")
    range_index = first_index(paragraphs, r"第八页\s*[-—至]\s*第九页")
    range_action = first_index(paragraphs, r"小标题.*对应内容.*不同动画")
    slide10_title = first_index(paragraphs, r"原来第十页删除.*Match five stages")
    slide10_asset = next(asset["paragraph"] for asset in assets if asset["paragraph"] > slide10_title)
    slide10_size = first_index(paragraphs, r"图片放到标题下方.*图宽")
    slide11 = first_index(paragraphs, r"第十一页.*相同的出现动画")
    removal = first_index(paragraphs, r"右下角图片删除")
    transitions = first_index(paragraphs, r"所有页面.*不同的切换动画")

    width3 = float(re.search(r"图宽(?:要)?\s*(\d+(?:\.\d+)?)", paragraphs[slide3]).group(1))
    width10 = float(re.search(r"图宽(?:要)?\s*(\d+(?:\.\d+)?)", paragraphs[slide10_size]).group(1))
    presentation = Presentation(source_pptx_path)
    expected_text = {
        "R01": paragraphs[cover],
        "R02": paragraphs[cover],
        "R03": paragraphs[slide3],
        "R04": paragraphs[slide4],
        "R05": "\n".join([paragraphs[slide5_title], paragraphs[slide5_action]]),
        "R06": paragraphs[slide6],
        "R07": "\n".join([paragraphs[range_index], paragraphs[range_action]]),
        "R08": "\n".join([paragraphs[range_index], paragraphs[range_action]]),
        "R09": "\n".join([paragraphs[slide10_title], paragraphs[slide10_size]]),
        "R10": paragraphs[slide11],
        "R11": paragraphs[removal],
        "R12": paragraphs[transitions],
    }
    expected_slides = {
        "R01": [1],
        "R02": [2],
        "R03": page_mentions(paragraphs[slide3]),
        "R04": page_mentions(paragraphs[slide4]),
        "R05": page_mentions(paragraphs[slide5_title]),
        "R06": page_mentions(paragraphs[slide6]),
        "R07": [page_mentions(paragraphs[range_index])[0]],
        "R08": [page_mentions(paragraphs[range_index])[1]],
        "R09": [10],
        "R10": page_mentions(paragraphs[slide11]),
        "R11": page_mentions(paragraphs[removal]),
        "R12": list(range(1, len(presentation.slides) + 1)),
    }
    return {
        "paragraphs": paragraphs,
        "assets": assets,
        "expected_text": expected_text,
        "expected_slides": expected_slides,
        "slide3_width": width3,
        "slide3_height": width3 / assets[0]["display_ratio"],
        "slide10_width": width10,
        "slide10_height": width10 / assets[3]["display_ratio"],
        "slide_count": len(presentation.slides),
        "requirements_sha256": sha256_file(requirements_path),
        "source_pptx_sha256": sha256_file(source_pptx_path),
    }


def requirement_map(plan: dict) -> dict[str, dict]:
    result: dict[str, dict] = {}
    for item in plan.get("requirements", []):
        identifier = item.get("id")
        if not isinstance(identifier, str) or identifier in result:
            raise ValueError("revision_plan.json 含缺失或重复 ID")
        result[identifier] = item
    return result


def validate_plan(plan: dict, truth: dict) -> tuple[bool, list[str]]:
    errors: list[str] = []
    try:
        mapped = requirement_map(plan)
    except Exception as error:
        return False, [str(error)]
    if set(mapped) != set(EXPECTED_ACTIONS):
        errors.append(f"要求 ID 不完整：{sorted(mapped)}")
        return False, errors
    if plan.get("requirements_source", {}).get("sha256") != truth["requirements_sha256"]:
        errors.append("要求文档 SHA256 与输入不一致")
    if plan.get("presentation_source", {}).get("sha256") != truth["source_pptx_sha256"]:
        errors.append("源 PPTX SHA256 与输入不一致")
    if plan.get("presentation_source", {}).get("slide_count") != truth["slide_count"]:
        errors.append("源 PPTX 页数记录错误")
    raw_paragraphs = [item.get("text") for item in plan.get("raw_paragraphs", [])]
    if raw_paragraphs != truth["paragraphs"]:
        errors.append("raw_paragraphs 未忠实反映要求.docx")
    output_assets = plan.get("embedded_assets", [])
    if [item.get("sha256") for item in output_assets] != [item["sha256"] for item in truth["assets"]]:
        errors.append("内嵌图片 SHA256 或文档顺序错误")
    for identifier, action in EXPECTED_ACTIONS.items():
        item = mapped[identifier]
        if item.get("action") != action:
            errors.append(f"{identifier} action 错误")
        if item.get("slides") != truth["expected_slides"][identifier]:
            errors.append(f"{identifier} 目标页错误：{item.get('slides')}")
        if item.get("source_text") != truth["expected_text"][identifier]:
            errors.append(f"{identifier} 未保留原始审阅意见")
    if not math.isclose(float(mapped["R03"].get("width_cm", -1)), truth["slide3_width"], abs_tol=1e-6):
        errors.append("R03 图片宽度未动态解析")
    if not math.isclose(
        float(mapped["R03"].get("height_cm_from_source_display_ratio", -1)),
        truth["slide3_height"],
        abs_tol=0.01,
    ):
        errors.append("R03 图片高度未从内嵌显示比例重算")
    if not math.isclose(float(mapped["R09"].get("width_cm", -1)), truth["slide10_width"], abs_tol=1e-6):
        errors.append("R09 图片宽度错误")
    if not math.isclose(
        float(mapped["R09"].get("height_cm_from_source_display_ratio", -1)),
        truth["slide10_height"],
        abs_tol=0.01,
    ):
        errors.append("R09 图片高度未从内嵌显示比例重算")
    return not errors, errors


def mutate_requirements(source: Path, destination: Path) -> None:
    with ZipFile(source, "r") as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    root = etree.fromstring(members["word/document.xml"])
    width_replacements = 0
    page_replacements = 0
    for node in root.xpath(".//w:t", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
        if node.text and "图宽12厘米" in node.text:
            node.text = node.text.replace("图宽12厘米", "图宽13厘米")
            width_replacements += 1
        if node.text and "第十六页右下角图片删除" in node.text:
            node.text = node.text.replace("第十六页右下角图片删除", "第十七页右下角图片删除")
            page_replacements += 1
    if width_replacements != 1 or page_replacements != 1:
        raise RuntimeError("无法构造内容变异要求文档")
    members["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for name, payload in members.items():
            archive.writestr(name, payload)


def run_processor(processor: Path, requirements: Path, source_pptx: Path, output: Path) -> tuple[bool, str]:
    command = [
        sys.executable,
        str(processor),
        "--requirements",
        str(requirements),
        "--source-pptx",
        str(source_pptx),
        "--output",
        str(output),
    ]
    completed = subprocess.run(command, capture_output=True, text=True, timeout=120)
    if completed.returncode != 0:
        return False, (completed.stderr or completed.stdout)[-2000:]
    return output.is_file(), "processor replay 完成"


def get_shape(slide, needle: str, *, exact: bool = False):
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = shape.text.strip()
        if (exact and text == needle) or (not exact and needle in text):
            return shape
    return None


def font_sizes(shape) -> list[float]:
    sizes: list[float] = []
    if shape is None:
        return sizes
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size is not None:
                sizes.append(run.font.size.pt)
    return sizes


def picture_hash(shape) -> str:
    return sha256_bytes(shape.image.blob)


def background_rgb(slide) -> str | None:
    background = slide._element.find(".//p:cSld/p:bg", NS)
    if background is None:
        return None
    color = background.find(".//a:srgbClr", NS)
    return color.get("val").upper() if color is not None else None


def center_x(shape) -> float:
    return shape.left + shape.width / 2


def overlap_fraction(first, second) -> float:
    left = max(first.left, second.left)
    top = max(first.top, second.top)
    right = min(first.left + first.width, second.left + second.width)
    bottom = min(first.top + first.height, second.top + second.height)
    if right <= left or bottom <= top:
        return 0.0
    intersection = (right - left) * (bottom - top)
    denominator = min(first.width * first.height, second.width * second.height)
    return intersection / denominator if denominator else 0.0


def effect_targets(slide) -> dict[str, set[str]]:
    timing = slide._element.find("p:timing", NS)
    result: dict[str, set[str]] = {}
    if timing is None:
        return result
    for effect in timing.findall(".//p:animEffect", NS):
        filter_name = effect.get("filter") or ""
        transition = effect.get("transition") or ""
        target = effect.find(".//p:spTgt", NS)
        behavior_time = effect.find(".//p:cBhvr/p:cTn", NS)
        if target is None or behavior_time is None:
            continue
        duration = behavior_time.get("dur") or "0"
        node = effect.getparent()
        preset_class = None
        while node is not None:
            if node.get("presetClass"):
                preset_class = node.get("presetClass")
                break
            node = node.getparent()
        valid = (
            transition == "in"
            and filter_name in APPROVED_EFFECTS
            and preset_class == "entr"
            and duration.isdigit()
            and int(duration) > 0
        )
        if valid:
            result.setdefault(target.get("spid"), set()).add(filter_name)
    return result


def click_targets(slide) -> list[str]:
    timing = slide._element.find("p:timing", NS)
    if timing is None:
        return []
    result: list[str] = []
    for time_node in timing.findall(".//p:cTn", NS):
        if time_node.get("nodeType") != "clickEffect":
            continue
        target = time_node.find(".//p:animEffect//p:spTgt", NS)
        if target is not None and target.get("spid"):
            result.append(target.get("spid"))
    return result


def group_animation_ok(slide, pairs: list[tuple[str, str]]) -> tuple[bool, str]:
    effects = effect_targets(slide)
    clicks = click_targets(slide)
    group_signatures: list[tuple[str, ...]] = []
    for heading, body in pairs:
        heading_shape = get_shape(slide, heading, exact=True)
        body_shape = get_shape(slide, body, exact=True)
        if heading_shape is None or body_shape is None:
            return False, f"缺少动画组 {heading!r}"
        heading_effect = effects.get(str(heading_shape.shape_id), set())
        body_effect = effects.get(str(body_shape.shape_id), set())
        if len(heading_effect) != 1 or heading_effect != body_effect:
            return False, f"{heading!r} 标题与正文动画不一致"
        if str(heading_shape.shape_id) not in clicks or str(body_shape.shape_id) not in clicks:
            return False, f"{heading!r} 未分别设置单击触发"
        group_signatures.append(tuple(sorted(heading_effect)))
    return len(set(group_signatures)) == 4, f"组动画={group_signatures}"


def transition_signature(slide) -> str | None:
    transition = slide._element.find("p:transition", NS)
    if transition is None or len(transition) != 1:
        return None
    child = transition[0]
    name = etree.QName(child).localname
    attributes = "|".join(f"{key}={value}" for key, value in sorted(child.attrib.items()))
    return f"{name}|{attributes}"


def validate_response(path: Path, truth: dict, deck_checks_passed: bool) -> tuple[bool, list[str]]:
    errors: list[str] = []
    workbook = load_workbook(path, data_only=False)
    if workbook.sheetnames != ["响应记录"]:
        errors.append(f"工作表必须且只能为“响应记录”：{workbook.sheetnames}")
        return False, errors
    sheet = workbook["响应记录"]
    if sheet.max_row != 13 or sheet.max_column != 5:
        errors.append(f"响应表范围必须为 A1:E13，实际 {sheet.max_row}x{sheet.max_column}")
        return False, errors
    headers = [sheet.cell(1, column).value for column in range(1, 6)]
    if headers != ["ID", "目标页", "原始审阅意见", "状态", "交叉文件核验"]:
        errors.append("响应表表头错误")
    seen: set[str] = set()
    for row in range(2, 14):
        identifier, target, source_text, status, check = [sheet.cell(row, column).value for column in range(1, 6)]
        if identifier not in EXPECTED_ACTIONS or identifier in seen:
            errors.append(f"第 {row} 行 ID 非法或重复：{identifier!r}")
            continue
        seen.add(identifier)
        expected_target = "1-18" if identifier == "R12" else ",".join(
            str(page) for page in truth["expected_slides"][identifier]
        )
        if str(target) != expected_target:
            errors.append(f"{identifier} 目标页与要求/PPT 冲突")
        if source_text != truth["expected_text"][identifier]:
            errors.append(f"{identifier} 原始审阅意见缺失或含无来源断言")
        if status != "已完成" or not deck_checks_passed:
            errors.append(f"{identifier} 状态与 PPT 实际完成情况冲突")
        if check != EXPECTED_RESPONSE_CHECKS[identifier]:
            errors.append(f"{identifier} 交叉文件核验字段错误或含无来源断言")
    if seen != set(EXPECTED_ACTIONS):
        errors.append("响应表遗漏审阅意见")
    return not errors, errors


def libreoffice_render_check(pptx: Path, workbook: Path) -> tuple[bool, str]:
    with tempfile.TemporaryDirectory(prefix="officeval-080-render-") as temporary:
        temporary_path = Path(temporary)
        profile = temporary_path / "profile"
        output = temporary_path / "pdf"
        output.mkdir()
        for source in [pptx, workbook]:
            command = [
                "libreoffice",
                f"-env:UserInstallation=file://{profile}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output),
                str(source),
            ]
            completed = subprocess.run(command, capture_output=True, text=True, timeout=180)
            if completed.returncode != 0:
                return False, (completed.stderr or completed.stdout)[-1000:]
        ppt_pdf = output / f"{pptx.stem}.pdf"
        xlsx_pdf = output / f"{workbook.stem}.pdf"
        if not ppt_pdf.is_file() or not xlsx_pdf.is_file():
            return False, "LibreOffice 未生成两个 PDF"
        info_ppt = subprocess.run(["pdfinfo", str(ppt_pdf)], capture_output=True, text=True, timeout=30)
        info_xlsx = subprocess.run(["pdfinfo", str(xlsx_pdf)], capture_output=True, text=True, timeout=30)
        ppt_pages = re.search(r"^Pages:\s+(\d+)", info_ppt.stdout, re.M)
        xlsx_pages = re.search(r"^Pages:\s+(\d+)", info_xlsx.stdout, re.M)
        valid = bool(ppt_pages and int(ppt_pages.group(1)) == 18 and xlsx_pages and int(xlsx_pages.group(1)) >= 1)
        return valid, f"PPT PDF={ppt_pages.group(1) if ppt_pages else None}页，XLSX PDF={xlsx_pages.group(1) if xlsx_pages else None}页"


def evaluate(artifact_dir: Path, input_dir: Path, *, run_render: bool = True) -> dict:
    checks: list[Check] = []
    missing = [name for name in REQUIRED_FILES if not (artifact_dir / name).is_file()]
    checks.append(Check("C01", "四个指定交付文件齐全", not missing, f"缺失={missing}", 4, True))
    if missing:
        return finalize(checks, metamorphic={"passed": False, "detail": "缺少交付文件"})

    source_pptx_path = input_dir / "公开课课件.pptx"
    requirements_path = input_dir / "要求.docx"
    try:
        truth = input_truth(requirements_path, source_pptx_path)
    except Exception as error:
        checks.append(Check("C02", "从真实输入独立重算真值", False, str(error), 4, True))
        return finalize(checks, metamorphic={"passed": False, "detail": str(error)})
    checks.append(Check("C02", "从真实输入独立重算真值", True, "12 条审阅要求与 4 张内嵌图片已重算", 4, True))

    try:
        supplied_plan = json.loads((artifact_dir / OUTPUT_PLAN).read_text(encoding="utf-8"))
        plan_ok, plan_errors = validate_plan(supplied_plan, truth)
    except Exception as error:
        plan_ok, plan_errors = False, [str(error)]
    checks.append(Check("C03", "revision_plan 忠实覆盖全部审阅意见和真实输入", plan_ok, "; ".join(plan_errors) or "计划与独立真值一致", 4, True))

    replay_detail: list[str] = []
    replay_ok = True
    with tempfile.TemporaryDirectory(prefix="officeval-080-replay-") as temporary:
        temp = Path(temporary)
        original_output = temp / "original.json"
        mutated_doc = temp / "要求_内容变异.docx"
        mutated_output = temp / "mutated.json"
        processor_path = artifact_dir / OUTPUT_PROCESSOR
        ok, detail = run_processor(processor_path, requirements_path, source_pptx_path, original_output)
        replay_detail.append(detail)
        replay_ok &= ok
        if ok:
            try:
                replay_plan = json.loads(original_output.read_text(encoding="utf-8"))
                same_plan, errors = validate_plan(replay_plan, truth)
                replay_ok &= same_plan
                replay_detail.extend(errors)
            except Exception as error:
                replay_ok = False
                replay_detail.append(str(error))
        try:
            mutate_requirements(requirements_path, mutated_doc)
            mutated_truth = input_truth(mutated_doc, source_pptx_path)
            ok, detail = run_processor(processor_path, mutated_doc, source_pptx_path, mutated_output)
            replay_detail.append(detail)
            replay_ok &= ok
            if ok:
                mutated_plan = json.loads(mutated_output.read_text(encoding="utf-8"))
                mutated_ok, errors = validate_plan(mutated_plan, mutated_truth)
                replay_ok &= mutated_ok
                replay_detail.extend(errors)
                mapped = requirement_map(mutated_plan)
                content_changed = (
                    math.isclose(float(mapped["R03"]["width_cm"]), 13.0, abs_tol=1e-6)
                    and mapped["R11"]["slides"] == [17]
                    and mapped["R03"]["source_text"].find("图宽13厘米") >= 0
                    and mapped["R11"]["source_text"].find("第十七页") >= 0
                )
                replay_ok &= content_changed
                if not content_changed:
                    replay_detail.append("处理器未随内容变异更新宽度、目标页和来源文字")
        except Exception as error:
            replay_ok = False
            replay_detail.append(str(error))
    checks.append(Check("C04", "处理器原输入 replay 与内容变异 replay 均通过", replay_ok, "; ".join(replay_detail) or "通过", 5, True))

    deck_path = artifact_dir / OUTPUT_PPTX
    source = Presentation(source_pptx_path)
    try:
        output = Presentation(deck_path)
        structural_ok = len(output.slides) == 18
        total_text = sum(
            1
            for slide in output.slides
            for shape in slide.shapes
            if getattr(shape, "has_text_frame", False) and shape.text.strip()
        )
        structural_ok &= total_text >= 120
        checks.append(Check("C05", "PPTX 可打开、18 页且保持原生可编辑对象", structural_ok, f"页数={len(output.slides)}，文本对象={total_text}", 5, True))
    except Exception as error:
        checks.append(Check("C05", "PPTX 可打开、18 页且保持原生可编辑对象", False, str(error), 5, True))
        return finalize(checks, metamorphic={"passed": replay_ok, "detail": replay_detail})

    cover_anchors = [
        "Comics & Animation",
        "Story Writing Lab",
        "Build a story from panels, bubbles, and action.",
        "observe visual clues",
    ]
    cover_errors: list[str] = []
    for anchor in cover_anchors:
        source_shape = get_shape(source.slides[0], anchor)
        output_shape = get_shape(output.slides[0], anchor)
        source_sizes = font_sizes(source_shape)
        output_sizes = font_sizes(output_shape)
        if not source_sizes or not output_sizes or len(output_sizes) != len(source_sizes):
            cover_errors.append(f"{anchor} 字号对象缺失")
            continue
        if not all(2.0 <= out - before <= 3.0 for before, out in zip(source_sizes, output_sizes)):
            cover_errors.append(f"{anchor} 未增加 2-3pt：{source_sizes}->{output_sizes}")
    bg2, bg3 = background_rgb(output.slides[1]), background_rgb(output.slides[2])
    background_ok = bg2 == bg3 and bg2 == "FFF7E8"
    checks.append(Check("C06", "第1页字体增加 2-3pt，第2页与第3页同为米黄色背景", not cover_errors and background_ok, "; ".join(cover_errors) or f"背景={bg2}", 4))

    slide3 = output.slides[2]
    student_hash = truth["assets"][0]["sha256"]
    student_pictures = [shape for shape in slide3.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and picture_hash(shape) == student_hash]
    slide3_ok = False
    slide3_detail = "未找到要求.docx 的学生图片"
    if len(student_pictures) == 1:
        picture = student_pictures[0]
        text_shapes = [shape for shape in slide3.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()]
        overlap = max((overlap_fraction(picture, shape) for shape in text_shapes), default=0)
        slide3_ok = (
            abs(picture.width / EMU_PER_CM - truth["slide3_width"]) <= 0.05
            and abs(picture.height / EMU_PER_CM - truth["slide3_height"]) <= 0.05
            and center_x(picture) > output.slide_width / 2
            and overlap < 0.05
        )
        slide3_detail = f"尺寸={picture.width/EMU_PER_CM:.2f}x{picture.height/EMU_PER_CM:.2f}cm，最大文字重叠={overlap:.3f}"
    checks.append(Check("C07", "第3页使用真实内嵌学生图片，动态尺寸正确且位于右侧", slide3_ok, slide3_detail, 4))

    slide4 = output.slides[3]
    source_slide4_picture = next(shape for shape in source.slides[3].shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE)
    slide4_pictures = [shape for shape in slide4.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and picture_hash(shape) == picture_hash(source_slide4_picture)]
    right_text = [get_shape(slide4, needle) for needle in ["Warm-up: meet the explorer", "What kind of story", "Who is the hero", "stronger plot"]]
    slide4_ok = len(slide4_pictures) == 1 and all(shape is not None for shape in right_text)
    if slide4_ok:
        picture = slide4_pictures[0]
        slide4_ok = center_x(picture) < output.slide_width / 2 and all(center_x(shape) > output.slide_width / 2 for shape in right_text) and picture.left + picture.width <= min(shape.left for shape in right_text)
    checks.append(Check("C08", "第4页保留真实原图并形成左图右文布局", slide4_ok, f"真实图片数={len(slide4_pictures)}", 4))

    slide5 = output.slides[4]
    pictures5 = [shape for shape in slide5.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    hashes5 = [picture_hash(shape) for shape in pictures5]
    expected5 = [truth["assets"][1]["sha256"], truth["assets"][2]["sha256"]]
    ratio_matches = sorted(round(shape.width / shape.height, 3) for shape in pictures5)
    target_ratios = sorted(round(truth["assets"][index]["display_ratio"], 3) for index in [1, 2])
    text5 = [shape for shape in slide5.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()]
    overlap5 = max((overlap_fraction(picture, text) for picture in pictures5 for text in text5), default=0)
    slide5_ok = sorted(hashes5) == sorted(expected5) and len(pictures5) == 2 and all(abs(a - b) <= 0.02 for a, b in zip(ratio_matches, target_ratios)) and overlap5 < 0.05
    checks.append(Check("C09", "第5页两图来自真实 Word 内嵌附件，显示比例正确且不遮字", slide5_ok, f"图片SHA匹配={sorted(hashes5)==sorted(expected5)}，比例={ratio_matches}，重叠={overlap5:.3f}", 4))

    slide6 = output.slides[5]
    arrows = [shape for shape in slide6.shapes if getattr(shape, "has_text_frame", False) and "→" in shape.text]
    old_answers = ["start the trip", "face danger", "make a choice", "find a way out"]
    arrow_ok = len(arrows) == 4
    for shape in arrows:
        right = shape.text.split("→", 1)[1]
        arrow_ok &= len(re.findall(r"_+", right)) == 2 and re.sub(r"_|\s", "", right) == ""
    arrow_ok &= not any(answer in "\n".join(shape.text for shape in slide6.shapes if getattr(shape, "has_text_frame", False)) for answer in old_answers)
    checks.append(Check("C10", "第6页每个箭头后仅保留两条横线", arrow_ok, f"箭头条目={len(arrows)}", 3))

    slide8_ok, slide8_detail = group_animation_ok(
        output.slides[7],
        [
            ("Speech bubble", "words spoken by characters"),
            ("Caption", "setting or scene information"),
            ("Sound effect", "noises in the scene"),
            ("Thought bubble", "words inside a character’s mind"),
        ],
    )
    slide9_ok, slide9_detail = group_animation_ok(
        output.slides[8],
        [
            ("Caption", "The scene begins at sunset."),
            ("Speech bubble", "“Watch out!”"),
            ("Thought bubble", "I hope the boat is strong."),
            ("Sound effect", "Crash!"),
        ],
    )
    checks.append(Check("C11", "第8-9页每组标题与正文分别可触发，四组效果互异", slide8_ok and slide9_ok, f"P8 {slide8_detail}; P9 {slide9_detail}", 5))

    slide10 = output.slides[9]
    title10 = get_shape(slide10, "Match five stages of the story with frames", exact=True)
    curve_hash = truth["assets"][3]["sha256"]
    curves = [shape for shape in slide10.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and picture_hash(shape) == curve_hash]
    slide10_ok = title10 is not None and len(curves) == 1
    if slide10_ok:
        curve = curves[0]
        slide10_ok = abs(curve.width / EMU_PER_CM - truth["slide10_width"]) <= 0.01 and abs(curve.height / EMU_PER_CM - truth["slide10_height"]) <= 0.01 and curve.top >= title10.top + title10.height
    checks.append(Check("C12", "第10页按真实意见替换标题并使用真实曲线图", slide10_ok, f"曲线图SHA匹配数={len(curves)}", 4))

    slide11 = output.slides[10]
    required11 = [get_shape(slide11, "Task 1:"), get_shape(slide11, "11", exact=True)]
    dialogues11 = [shape for shape in slide11.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip().startswith("“")]
    dialogues11.sort(key=lambda shape: (shape.top, shape.left))
    required11.extend(dialogues11)
    effects11 = effect_targets(slide11)
    signatures11 = [tuple(sorted(effects11.get(str(shape.shape_id), set()))) for shape in required11 if shape is not None]
    clicks11 = click_targets(slide11)
    ids11 = [str(shape.shape_id) for shape in required11 if shape is not None]
    slide11_ok = len(required11) == 8 and all(shape is not None for shape in required11) and len(set(signatures11)) == 1 and signatures11 and signatures11[0] == ("fade",) and clicks11 == ids11
    checks.append(Check("C13", "第11页副标题、页码和六条对白同效果且按顺序分别触发", slide11_ok, f"对象={len(required11)}，点击序列={clicks11}", 5))

    slide16 = output.slides[15]
    lower_right_pictures = [shape for shape in slide16.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and center_x(shape) > output.slide_width / 2 and shape.top + shape.height / 2 > output.slide_height / 2]
    slide16_ok = len(lower_right_pictures) == 0
    transition_signatures = [transition_signature(slide) for slide in output.slides]
    transitions_ok = None not in transition_signatures and len(set(transition_signatures)) == 18
    checks.append(Check("C14", "第16页右下角图片已删除，18 页切换组合全部唯一", slide16_ok and transitions_ok, f"右下角图片={len(lower_right_pictures)}，唯一切换={len(set(transition_signatures))}/18", 5))

    preserved_errors: list[str] = []
    for page in list(range(1, 10)) + list(range(11, 19)):
        source_title = next((shape.text.strip() for shape in source.slides[page - 1].shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()), None)
        if source_title and get_shape(output.slides[page - 1], source_title, exact=True) is None:
            preserved_errors.append(f"P{page} 标题缺失")
    out_of_bounds = 0
    overlaps = 0
    for slide in output.slides:
        text_shapes = [shape for shape in slide.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()]
        for shape in text_shapes:
            if shape.left < 0 or shape.top < 0 or shape.left + shape.width > output.slide_width or shape.top + shape.height > output.slide_height:
                out_of_bounds += 1
        for index, first in enumerate(text_shapes):
            for second in text_shapes[index + 1 :]:
                if overlap_fraction(first, second) > 0.20:
                    overlaps += 1
    quality_ok = not preserved_errors and out_of_bounds == 0 and overlaps < 3
    checks.append(Check("C15", "未改页标题保留，文字不越界且无三处以上显著重叠", quality_ok, f"保留错误={preserved_errors}，越界={out_of_bounds}，显著重叠={overlaps}", 4))

    deck_checks_passed = all(check.passed for check in checks if check.id in {"C05", "C06", "C07", "C08", "C09", "C10", "C11", "C12", "C13", "C14", "C15"})
    try:
        response_ok, response_errors = validate_response(artifact_dir / OUTPUT_XLSX, truth, deck_checks_passed)
    except Exception as error:
        response_ok, response_errors = False, [str(error)]
    checks.append(Check("C16", "修订响应逐条覆盖真实意见并与 PPT 交叉一致，无无来源断言", response_ok, "; ".join(response_errors) or "12 条响应与 PPT 一致", 5, True))

    upstream_path = Path(__file__).resolve().parents[1] / "adaptation" / "upstream-verifier.py"
    if not upstream_path.is_file():
        upstream_path = Path("/app/upstream-verifier.py")
    upstream_ok = False
    upstream_detail = ""
    try:
        specification = importlib.util.spec_from_file_location("upstream_080", upstream_path)
        module = importlib.util.module_from_spec(specification)
        specification.loader.exec_module(module)
        upstream_result = module.evaluate(str(artifact_dir))
        upstream_ok = upstream_result.get("total_score") == 31 and upstream_result.get("max_score") == 31
        upstream_detail = f"{upstream_result.get('total_score')}/{upstream_result.get('max_score')}"
    except Exception as error:
        upstream_detail = str(error)
    checks.append(Check("C17", "保留的上游 verifier 达到 31/31", upstream_ok, upstream_detail, 3))

    if run_render:
        try:
            render_ok, render_detail = libreoffice_render_check(deck_path, artifact_dir / OUTPUT_XLSX)
        except Exception as error:
            render_ok, render_detail = False, str(error)
        checks.append(Check("C18", "PPTX 与 XLSX 均可由 LibreOffice 实际渲染", render_ok, render_detail, 4, True))

    return finalize(checks, metamorphic={"passed": replay_ok, "detail": replay_detail})


def finalize(checks: list[Check], metamorphic: dict) -> dict:
    total_weight = sum(check.weight for check in checks)
    earned = sum(check.weight for check in checks if check.passed)
    hurdle_failure = any(check.hurdle and not check.passed for check in checks)
    score = 0.0 if hurdle_failure else (earned / total_weight if total_weight else 0.0)
    return {
        "task_id": "cowork-evo/officeval-080-requirements-driven-deck-revision",
        "score": round(score, 6),
        "passed": math.isclose(score, 1.0, abs_tol=1e-9),
        "checks": [asdict(check) for check in checks],
        "metamorphic_replay": metamorphic,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--skip-render", action="store_true")
    arguments = parser.parse_args()
    result = evaluate(arguments.artifact_dir, arguments.input_dir, run_render=not arguments.skip_render)
    arguments.result.parent.mkdir(parents=True, exist_ok=True)
    arguments.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
