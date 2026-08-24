#!/usr/bin/env python3
"""officeval_076 大纲到动画演示文稿任务的 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN

import upstream_verifier


OUTPUT_NAME = "艺术漆北区销售额提升方案.pptx"
TASK_ID = "cowork-evo/officeval-076-outline-to-animated-sales-plan"
RUBRIC_VERSION = "omegause-officeval-076-adaptation-v0.1"
JUDGE_VERSION = "officeval-076-adaptation-v0.1"
PASS_THRESHOLD = 0.8
TEMPLATE_BACKGROUND_SHA256 = "9112d4e712200684e80dc40d0dea06c6e5747f43f08e41b1bf7560331eb729c6"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}
BODY_PAGES = [4, 6, 8, 10, 12, 14, 16, 18, 20]
SECTION_PAGES = [3, 5, 7, 9, 11, 13, 15, 17, 19]

CRITERION_SPECS = {
    "R001": ("交付指定名称的可打开 21 页 PPTX", 2, "hurdle"),
    "R002": ("封面、目录、九个板块、十八个二级标题与原始正文完整", 5, "hurdle"),
    "R003": ("封面与第 1–20 页浅绿到中绿渐变正确", 4, "positive"),
    "R004": ("目录页与九张章节页的页序、文本和层级正确", 5, "positive"),
    "R005": ("九张正文页的标题、原文、字号和上下顺序正确", 5, "positive"),
    "R006": ("九张正文页均有合规橙色可编辑外框", 3, "positive"),
    "R007": ("所有可见中文都使用宋体或 SimSun", 3, "positive"),
    "R008": ("十八个二级标题和十八段正文都有区分明显的入场动画", 5, "positive"),
    "R009": ("九张正文页均为一次单击后按顺序自动连续播放", 5, "positive"),
    "R010": ("第 21 页使用真实模板背景并只保留相关致谢文字", 3, "positive"),
    "R011": ("LibreOffice 实际渲染为 21 页且关键文字可读", 2, "positive"),
    "R012": ("前 20 页的文本和外框为原生可编辑对象", 3, "positive"),
    "R013": ("18 页以上缺少规定的浅绿渐变", -8, "penalty"),
    "R014": ("正文页动画主序列缺失、严重不全或需要多次点击", -10, "penalty"),
    "R015": ("缺失真实模板尾页、前 20 页被压平或广泛使用非宋体", -8, "penalty"),
}


def criterion(
    identifier: str,
    passed: bool,
    evidence: str = "",
    score: float | None = None,
) -> dict:
    description, weight, kind = CRITERION_SPECS[identifier]
    if score is None:
        score = 1.0 if passed else 0.0
    return {
        "criterion_id": identifier,
        "description": description,
        "weight": weight,
        "type": kind,
        "passed": bool(passed),
        "score": max(0.0, min(1.0, float(score))),
        "evidence": evidence,
        "method": "deterministic",
    }


def shape_text(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    return shape.text_frame.text.strip()


def visible_runs(shape):
    if not getattr(shape, "has_text_frame", False):
        return []
    return [
        run
        for paragraph in shape.text_frame.paragraphs
        for run in paragraph.runs
        if (run.text or "").strip()
    ]


def exact_shape(slide, text: str):
    matches = [shape for shape in slide.shapes if shape_text(shape) == text]
    return matches[0] if len(matches) == 1 else None


def load_outline() -> dict:
    path = Path(__file__).resolve().parents[1] / "solution" / "outline.json"
    return json.loads(path.read_text(encoding="utf-8"))


def check_exact_content(prs: Presentation, outline: dict) -> tuple[bool, str]:
    errors = []
    if exact_shape(prs.slides[0], outline["title"]) is None:
        errors.append("封面标题缺失或重复")
    toc_text = "\n".join(shape_text(shape) for shape in prs.slides[1].shapes)
    if outline["toc_title"] not in toc_text:
        errors.append("第 2 页缺少目录标题")

    for index, section in enumerate(outline["sections"]):
        section_page = SECTION_PAGES[index]
        body_page = BODY_PAGES[index]
        if exact_shape(prs.slides[section_page - 1], section["section_title"]) is None:
            errors.append(f"P{section_page} 缺少章节标题")
        if section["toc_title"] not in toc_text:
            errors.append(f"目录缺少 {section['toc_title']}")
        body_slide = prs.slides[body_page - 1]
        for subsection in section["subsections"]:
            if exact_shape(body_slide, subsection["title"]) is None:
                errors.append(f"P{body_page} 缺少 {subsection['title']}")
            if exact_shape(body_slide, subsection["body"]) is None:
                errors.append(f"P{body_page} 缺少或改写 {subsection['title']} 正文")

    closing_text = "\n".join(shape_text(shape) for shape in prs.slides[20].shapes)
    for value in outline["closing"].values():
        if value not in closing_text:
            errors.append(f"尾页缺少 {value}")
    return not errors, "文本全部精确命中" if not errors else "；".join(errors[:20])


def check_body_pages_corrected(prs: Presentation, outline: dict) -> tuple[bool, str]:
    """使用原始 Query 精确正文修正上游 3.1 关键词错误。"""
    errors = []
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    for index, section in enumerate(outline["sections"]):
        page_no = BODY_PAGES[index]
        slide = prs.slides[page_no - 1]
        title_shapes = []
        for subsection_index, subsection in enumerate(section["subsections"]):
            title = exact_shape(slide, subsection["title"])
            body = exact_shape(slide, subsection["body"])
            if title is None or body is None:
                errors.append(f"P{page_no} 第{subsection_index + 1}节文本不完整")
                continue
            title_shapes.append(title)
            title_runs = visible_runs(title)
            body_runs = visible_runs(body)
            left_ratio = (title.left or 0) / slide_width
            center_y = ((title.top or 0) + (title.height or 0) / 2) / slide_height
            expected_y = center_y <= 0.55 if subsection_index == 0 else 0.4 <= center_y <= 0.8
            alignment = title.text_frame.paragraphs[0].alignment
            title_ok = (
                left_ratio <= 0.3
                and expected_y
                and alignment in (None, PP_ALIGN.LEFT, PP_ALIGN.JUSTIFY)
                and title_runs
                and all(run.font.bold is True for run in title_runs)
                and all(run.font.size and 16 <= run.font.size.pt <= 20 for run in title_runs)
            )
            body_ok = (
                (body.top or 0) > (title.top or 0)
                and abs((body.left or 0) - (title.left or 0)) <= 0.1 * slide_width
                and body_runs
                and all(run.font.size and 12 <= run.font.size.pt <= 16 for run in body_runs)
            )
            if not title_ok:
                errors.append(f"P{page_no} {subsection['title']} 位置或字号不合规")
            if not body_ok:
                errors.append(f"P{page_no} {subsection['title']} 正文位置或字号不合规")
        if len(title_shapes) == 2 and not ((title_shapes[0].top or 0) < (title_shapes[1].top or 0)):
            errors.append(f"P{page_no} 两个小节顺序错误")
    return not errors, "9/9 正文页原文、位置与字号正确" if not errors else "；".join(errors[:20])


def safe_upstream(function, *arguments) -> tuple[bool, str]:
    try:
        ok, evidence = function(*arguments)
        return bool(ok), str(evidence)
    except Exception as exc:
        return False, f"上游检查异常：{exc}"


def gradient_page_count(raw_zip: ZipFile, page_count: int) -> int:
    count = 0

    def luminance(color: tuple[int, int, int]) -> float:
        red, green, blue = [value / 255 for value in color]
        return 0.2126 * red + 0.7152 * green + 0.0722 * blue

    for page_no in range(1, min(page_count, 20) + 1):
        try:
            root = etree.fromstring(raw_zip.read(f"ppt/slides/slide{page_no}.xml"))
        except Exception:
            continue
        colors = []
        for gradient in root.findall(".//a:gradFill", NS):
            for color in gradient.findall(".//a:srgbClr", NS):
                value = color.get("val")
                if value and re.fullmatch(r"[0-9A-Fa-f]{6}", value):
                    colors.append(upstream_verifier.hex_to_rgb(value))
            if colors:
                break
        if len(colors) < 2 or not all(upstream_verifier.is_green(color) for color in colors):
            continue
        levels = [luminance(color) for color in colors]
        if min(levels) >= 0.35 and max(levels) >= 0.75 and sum(levels) / len(levels) >= 0.55:
            count += 1
    return count


def count_bad_chinese_fonts(prs: Presentation) -> int:
    bad = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            for run in visible_runs(shape):
                if re.search(r"[\u3400-\u4dbf\u4e00-\u9fff]", run.text or ""):
                    if run.font.name not in ("宋体", "SimSun"):
                        bad += 1
    return bad


def template_background_present(prs: Presentation) -> tuple[bool, str]:
    hashes = []
    for shape in prs.slides[20].shapes:
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        digest = hashlib.sha256(shape.image.blob).hexdigest()
        hashes.append(digest)
    found = TEMPLATE_BACKGROUND_SHA256 in hashes
    return found, f"尾页图片哈希={hashes}"


def native_editability(prs: Presentation, outline: dict) -> tuple[bool, str, int]:
    errors = []
    full_slide_pictures = 0
    slide_area = prs.slide_width * prs.slide_height
    for page_no, slide in enumerate(list(prs.slides)[:20], 1):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                area = int(shape.width or 0) * int(shape.height or 0)
                if area >= 0.9 * slide_area:
                    full_slide_pictures += 1
                    errors.append(f"P{page_no} 出现整页图片")
    for index, section in enumerate(outline["sections"]):
        slide = prs.slides[BODY_PAGES[index] - 1]
        exact_texts = [
            section["subsections"][0]["title"],
            section["subsections"][0]["body"],
            section["subsections"][1]["title"],
            section["subsections"][1]["body"],
        ]
        if any(exact_shape(slide, text) is None for text in exact_texts):
            errors.append(f"P{BODY_PAGES[index]} 文本未保持原生可编辑")
        editable_boxes = [
            shape
            for shape in slide.shapes
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
        ]
        if len(editable_boxes) < 3:
            errors.append(f"P{BODY_PAGES[index]} 可编辑背景/外框数量不足")
    return not errors, "前 20 页未压平，正文与外框可编辑" if not errors else "；".join(errors[:20]), full_slide_pictures


def render_check(path: Path, outline: dict) -> tuple[bool, str]:
    libreoffice = shutil.which("libreoffice")
    pdfinfo = shutil.which("pdfinfo")
    pdftotext = shutil.which("pdftotext")
    if not libreoffice or not pdfinfo or not pdftotext:
        return False, "缺少 libreoffice、pdfinfo 或 pdftotext"

    with tempfile.TemporaryDirectory(prefix="cowork-evo-076-") as temp_dir_text:
        temp_dir = Path(temp_dir_text)
        home = temp_dir / "home"
        home.mkdir()
        environment = os.environ.copy()
        environment["HOME"] = str(home)
        for key in (
            "http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY",
            "ALL_PROXY", "all_proxy",
        ):
            environment.pop(key, None)
        converted = subprocess.run(
            [
                libreoffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(temp_dir),
                str(path),
            ],
            capture_output=True,
            text=True,
            timeout=240,
            check=False,
            env=environment,
        )
        pdf_path = temp_dir / f"{path.stem}.pdf"
        if converted.returncode != 0 or not pdf_path.is_file():
            return False, (converted.stderr or converted.stdout or "转换失败")[-1000:]
        info = subprocess.run(
            [pdfinfo, str(pdf_path)],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        extracted = subprocess.run(
            [pdftotext, "-layout", str(pdf_path), "-"],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        page_match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.M)
        pages = int(page_match.group(1)) if page_match else 0
        normalized = re.sub(r"\s+", "", extracted.stdout)
        anchors = [outline["title"], outline["toc_title"], outline["closing"]["thanks"]]
        anchors.extend(section["section_title"] for section in outline["sections"])
        anchors.extend(
            subsection["title"]
            for section in outline["sections"]
            for subsection in section["subsections"]
        )
        missing = [text for text in anchors if re.sub(r"\s+", "", text) not in normalized]
    ok = pages == 21 and extracted.returncode == 0 and not missing
    return ok, f"渲染页数={pages}，缺失关键文字={missing[:10]}"


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)
    try:
        presentation = Presentation(output_path)
        raw_zip = ZipFile(output_path)
    except Exception as exc:
        results.append(criterion("R001", False, f"文件无法打开：{exc}"))
        return build_result(results, output_path)

    width = presentation.slide_width / 914400
    height = presentation.slide_height / 914400
    integrity_ok = (
        len(presentation.slides) == 21
        and abs(width - 13.333333) <= 0.02
        and abs(height - 7.5) <= 0.02
    )
    results.append(
        criterion(
            "R001",
            integrity_ok,
            f"页数={len(presentation.slides)}，尺寸={width:.4f}×{height:.4f}英寸",
        )
    )
    if not integrity_ok:
        raw_zip.close()
        return build_result(results, output_path)

    outline = load_outline()
    content_ok, content_evidence = check_exact_content(presentation, outline)
    results.append(criterion("R002", content_ok, content_evidence))
    if not content_ok:
        raw_zip.close()
        return build_result(results, output_path)

    cover_ok, cover_evidence = safe_upstream(
        upstream_verifier.check_first_page,
        presentation,
        raw_zip,
    )
    background_ok, background_evidence = safe_upstream(
        upstream_verifier.check_bg_gradient,
        presentation,
        raw_zip,
    )
    results.append(
        criterion(
            "R003",
            cover_ok and background_ok,
            f"封面={cover_evidence}；背景={background_evidence}",
        )
    )

    toc_ok, toc_evidence = safe_upstream(upstream_verifier.check_toc, presentation)
    section_ok, section_evidence = safe_upstream(
        upstream_verifier.check_section_pages,
        presentation,
    )
    results.append(
        criterion(
            "R004",
            toc_ok and section_ok,
            f"目录={toc_evidence}；章节={section_evidence}",
        )
    )

    body_ok, body_evidence = check_body_pages_corrected(presentation, outline)
    results.append(criterion("R005", body_ok, body_evidence))
    orange_ok, orange_evidence = safe_upstream(
        upstream_verifier.check_orange_box,
        presentation,
    )
    results.append(criterion("R006", orange_ok, orange_evidence))
    font_ok, font_evidence = safe_upstream(
        upstream_verifier.check_all_simsun,
        presentation,
    )
    results.append(criterion("R007", font_ok, font_evidence))

    subtitle_anim_ok, subtitle_anim_evidence = safe_upstream(
        upstream_verifier.check_sub_title_anim,
        presentation,
        raw_zip,
    )
    body_anim_ok, body_anim_evidence = safe_upstream(
        upstream_verifier.check_body_anim,
        presentation,
        raw_zip,
    )
    effect_ok, effect_evidence = safe_upstream(
        upstream_verifier.check_anim_effect_diff,
        presentation,
        raw_zip,
    )
    animation_targets_ok = subtitle_anim_ok and body_anim_ok and effect_ok
    results.append(
        criterion(
            "R008",
            animation_targets_ok,
            f"标题={subtitle_anim_evidence}；正文={body_anim_evidence}；差异={effect_evidence}",
        )
    )

    sequence_checks = [
        safe_upstream(upstream_verifier.check_anim_trigger, presentation, raw_zip),
        safe_upstream(upstream_verifier.check_anim_order, presentation, raw_zip),
        safe_upstream(upstream_verifier.check_anim_continuity, presentation, raw_zip),
        safe_upstream(upstream_verifier.check_page_switch_logic, presentation, raw_zip),
    ]
    sequence_ok = all(item[0] for item in sequence_checks)
    results.append(
        criterion(
            "R009",
            sequence_ok,
            "；".join(item[1] for item in sequence_checks),
        )
    )

    last_ok, last_evidence = safe_upstream(
        upstream_verifier.check_last_page,
        presentation,
    )
    template_ok, template_evidence = template_background_present(presentation)
    results.append(
        criterion(
            "R010",
            last_ok and template_ok,
            f"尾页={last_evidence}；模板={template_evidence}",
        )
    )

    rendered, render_evidence = render_check(output_path, outline)
    results.append(criterion("R011", rendered, render_evidence))
    editable, editable_evidence, full_slide_pictures = native_editability(
        presentation,
        outline,
    )
    results.append(criterion("R012", editable, editable_evidence))

    gradient_pages = gradient_page_count(raw_zip, len(presentation.slides))
    results.append(
        criterion(
            "R013",
            gradient_pages < 18,
            f"合规渐变页数={gradient_pages}/20",
        )
    )
    results.append(
        criterion(
            "R014",
            not (animation_targets_ok and sequence_ok),
            "动画主序列完整" if animation_targets_ok and sequence_ok else "动画目标或序列不完整",
        )
    )
    bad_fonts = count_bad_chinese_fonts(presentation)
    severe_delivery_problem = (not template_ok) or full_slide_pictures > 0 or bad_fonts >= 5
    results.append(
        criterion(
            "R015",
            severe_delivery_problem,
            f"模板背景={template_ok}，前20页整页图={full_slide_pictures}，非宋体中文={bad_fonts}",
        )
    )
    raw_zip.close()
    return build_result(results, output_path)


def build_result(results: list[dict], artifact_path: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            results.append(
                criterion(
                    identifier,
                    False,
                    "因交付文件不存在、损坏或门槛失败，本判据未执行",
                )
            )
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(
        item["weight"] * item["score"]
        for item in results
        if item["weight"] > 0
    )
    penalties = sum(
        item["weight"]
        for item in results
        if item["weight"] < 0 and item["passed"]
    )
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [
        item["criterion_id"]
        for item in results
        if item["type"] == "hurdle" and not item["passed"]
    ]
    if hurdle_failures:
        reward = 0.0
    digest = (
        hashlib.sha256(artifact_path.read_bytes()).hexdigest()
        if artifact_path.is_file()
        else hashlib.sha256(b"").hexdigest()
    )
    normalized = []
    for item in results:
        normalized_score = (
            0.0 if item["passed"] else 1.0
        ) if item["type"] == "penalty" else item["score"]
        normalized.append(
            {
                "criterion_id": item["criterion_id"],
                "score": normalized_score,
                "evidence": item["evidence"],
                "method": item["method"],
                "raw": {
                    "description": item["description"],
                    "weight": item["weight"],
                    "type": item["type"],
                    "passed": item["passed"],
                },
            }
        )
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [digest],
        "criterion_results": normalized,
        "reward": round(reward, 6),
        "passed": reward >= PASS_THRESHOLD,
        "hurdle_failures": hurdle_failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    arguments = parser.parse_args()
    result = evaluate(arguments.artifact_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if arguments.result:
        arguments.result.parent.mkdir(parents=True, exist_ok=True)
        arguments.result.write_text(rendered + "\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
