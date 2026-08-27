#!/usr/bin/env python3
"""officeval_074 本地化任务的 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import upstream_verifier


OUTPUT_NAME = "农业县低空经济发展方案_English.pptx"
SOURCE_NAME = "农业县低空经济发展方案.pptx"
TASK_ID = "cowork-evo/officeval-074-full-deck-localization"
RUBRIC_VERSION = "omegause-officeval-074-adaptation-v0.1"
JUDGE_VERSION = "officeval-074-adaptation-v0.1"
PASS_THRESHOLD = 0.8

CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")
NON_ENGLISH_SCRIPT_RE = re.compile(
    r"[\u3400-\u4dbf\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]"
)
ENGLISH_WORD_RE = re.compile(r"[A-Za-z][A-Za-z'\-]{1,}")
EMU_PER_INCH = 914400

CRITERION_SPECS = {
    "R001": ("指定 PPTX 可打开且严格为 46 页", 2, "hurdle"),
    "R002": ("507 个源中文对象全部英文化且无非英文主要脚本残留", 8, "hurdle"),
    "R003": ("页序、对象结构、图片、背景和主体版式保持", 7, "hurdle"),
    "R004": ("46 页核心主题和业务术语语义正确", 8, "positive"),
    "R005": ("字体、加粗、填充和线条样式保持", 3, "positive"),
    "R006": ("满足上游封面、目录、章节页和页脚字号规则", 3, "positive"),
    "R007": ("LibreOffice 实际渲染为 46 页且文字可读", 2, "positive"),
    "R008": ("对象不越界且显式字号达到 7.8pt", 2, "positive"),
    "R009": ("残留中文、删除文字对象、丢页或文件损坏", -10, "penalty"),
    "R010": ("图片、背景、非文本对象或主体版式被改变", -10, "penalty"),
}

# 每页只放任务语义的最低锚点，不要求与 Golden 逐字一致；同一元组内是可接受同义写法。
SEMANTIC_ANCHORS = {
    1: [("qinglan",), ("agricultur",), ("low-altitude", "low altitude")],
    2: [("contents", "table of contents"), ("part 07",)],
    3: [("project background",), ("terrain", "agricultural labor")],
    4: [("low-altitude", "low altitude"), ("terrain",)],
    5: [("policy",), ("regional",)],
    6: [("agricultural drone", "agriculture drone"), ("scenarios",)],
    7: [("crop protection",), ("material transport", "transport")],
    8: [("traditional operations", "traditional operation"), ("aerial coordination",)],
    9: [("digital operating workflow", "digital workflow"), ("assign", "dispatch")],
    10: [("precision agriculture",), ("zone-specific", "zone specific", "management zones")],
    11: [("scenario integration",), ("county platform",)],
    12: [("rice-field", "rice field"), ("variable-rate", "variable rate")],
    13: [("cash-crop", "cash crop"), ("monitoring",)],
    14: [("economic",), ("social benefits", "social benefit")],
    15: [("cost",), ("efficiency",)],
    16: [("traditional method",), ("low-altitude method", "low altitude method")],
    17: [("key challenges", "key challenge"), ("adoption",)],
    18: [("pain points", "pain point"), ("training", "maintenance", "o&m")],
    19: [("core solutions", "core solution"), ("infrastructure",)],
    20: [("infrastructure",), ("county data platform", "data platform")],
    21: [("training",), ("o&m", "maintenance")],
    22: [("service procurement",), ("funding", "investment")],
    23: [("service procurement",), ("operating scenarios", "operating scenario")],
    24: [("benefit review", "benefits"), ("scaling", "scale")],
    25: [("wheat",), ("weeding", "weed")],
    26: [("citrus", "orchard"), ("pest-control", "pest control")],
    27: [("cotton",), ("defoliant",)],
    28: [("tea plantation", "tea"), ("aphid",)],
    29: [("medicinal crops", "medicinal crop"), ("aphid",)],
    30: [("peach-orchard", "peach orchard"), ("evaluation", "review")],
    31: [("disaster assessment",), ("insurance", "claims")],
    32: [("straw-burning", "straw burning"), ("patrol",)],
    33: [("transport task",), ("route planning", "flight plan")],
    34: [("takeoff",), ("delivery", "return")],
    35: [("parcel verification",), ("air-ground", "air–ground", "aerial collection")],
    36: [("comparison", "compare"), ("verification methods", "verification method")],
    37: [("imagery", "image"), ("conclusion", "workflow")],
    38: [("crop-protection", "crop protection"), ("configuration",)],
    39: [("transport drone",), ("configuration",)],
    40: [("inspection drone",), ("configuration",)],
    41: [("mobile reshoot",), ("inspection",)],
    42: [("summary",), ("outlook",)],
    43: [("agricultural product",), ("oversight", "governance")],
    44: [("platform architecture",), ("regulator", "regulatory")],
    45: [("work plan",), ("expansion",)],
    46: [("thank you", "thanks"), ("qinglan county", "qinglan")],
}

# 这些对象的源文本框本来就只有 0.01–0.18 英寸高，本地参考解只在原卡片内修复可读性。
TEXT_FIT_EXCEPTIONS = {
    (13, 16),
    (20, 16),
    (26, 16),
    (27, 16),
    (30, 16),
    (31, 16),
    (43, 16),
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


def normalized_text(text: str) -> str:
    value = (text or "").lower()
    value = value.replace("–", "-").replace("—", "-").replace("‑", "-")
    return re.sub(r"\s+", " ", value).strip()


def slide_text(slide) -> str:
    return "\n".join(
        shape.text_frame.text
        for shape in slide.shapes
        if getattr(shape, "has_text_frame", False)
    )


def visible_runs(shape):
    if not getattr(shape, "has_text_frame", False):
        return []
    return [
        run
        for paragraph in shape.text_frame.paragraphs
        for run in paragraph.runs
        if (run.text or "").strip()
    ]


def color_signature(color) -> tuple:
    try:
        color_type = str(color.type)
    except Exception:
        color_type = None
    try:
        rgb = str(color.rgb) if color.rgb is not None else None
    except Exception:
        rgb = None
    try:
        theme = str(color.theme_color) if color.theme_color is not None else None
    except Exception:
        theme = None
    try:
        brightness = round(float(color.brightness), 5)
    except Exception:
        brightness = None
    return color_type, rgb, theme, brightness


def shape_style_signature(shape) -> tuple:
    try:
        auto_shape_type = str(shape.auto_shape_type)
    except Exception:
        auto_shape_type = None
    try:
        fill_type = str(shape.fill.type)
        fill_fore = color_signature(shape.fill.fore_color)
        fill_back = color_signature(shape.fill.back_color)
    except Exception:
        fill_type, fill_fore, fill_back = None, None, None
    try:
        line_type = str(shape.line.fill.type)
        line_color = color_signature(shape.line.color)
        line_width = shape.line.width
        line_dash = str(shape.line.dash_style)
    except Exception:
        line_type, line_color, line_width, line_dash = None, None, None, None
    return (
        auto_shape_type,
        round(float(shape.rotation or 0), 4),
        fill_type,
        fill_fore,
        fill_back,
        line_type,
        line_color,
        line_width,
        line_dash,
    )


def background_signature(slide) -> bytes:
    background = slide.background._cSld.bg
    if background is None:
        return b""
    return etree.tostring(background, method="c14n")


def geometry_delta_inches(source_shape, candidate_shape) -> float:
    return max(
        abs(source_shape.left - candidate_shape.left),
        abs(source_shape.top - candidate_shape.top),
        abs(source_shape.width - candidate_shape.width),
        abs(source_shape.height - candidate_shape.height),
    ) / EMU_PER_INCH


def image_digest(shape) -> str | None:
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return None
    return hashlib.sha256(shape.image.blob).hexdigest()


def english_completion_check(source: Presentation, candidate: Presentation) -> tuple[bool, str]:
    expected = 0
    translated = 0
    failures = []
    for page, (source_slide, candidate_slide) in enumerate(
        zip(source.slides, candidate.slides), start=1
    ):
        for index, source_shape in enumerate(source_slide.shapes):
            source_value = getattr(source_shape, "text", "") or ""
            if not CJK_RE.search(source_value):
                continue
            expected += 1
            if index >= len(candidate_slide.shapes):
                failures.append(f"p{page}/shape{index}:对象缺失")
                continue
            candidate_value = getattr(candidate_slide.shapes[index], "text", "") or ""
            if not candidate_value.strip():
                failures.append(f"p{page}/shape{index}:英文为空")
                continue
            if NON_ENGLISH_SCRIPT_RE.search(candidate_value):
                failures.append(f"p{page}/shape{index}:仍含中日韩文字")
                continue
            if not ENGLISH_WORD_RE.search(candidate_value):
                failures.append(f"p{page}/shape{index}:没有可识别英文单词")
                continue
            translated += 1

    candidate_all = "\n".join(slide_text(slide) for slide in candidate.slides)
    extra_script = bool(NON_ENGLISH_SCRIPT_RE.search(candidate_all))
    upstream_p2, upstream_p2_detail = upstream_verifier.rule_p2_toc_all_english(candidate)
    upstream_p3, upstream_p3_detail = upstream_verifier.rule_p3_to_p46_all_english(candidate)
    passed = (
        expected == 507
        and translated == expected
        and not failures
        and not extra_script
        and upstream_p2
        and upstream_p3
    )
    evidence = (
        f"源中文对象={expected}，有效英文对象={translated}，"
        f"额外非英文脚本={extra_script}，目录英文={upstream_p2}，正文英文={upstream_p3}"
    )
    if failures:
        evidence += f"，首批错误={failures[:8]}"
    if not upstream_p2:
        evidence += f"，目录详情={upstream_p2_detail}"
    if not upstream_p3:
        evidence += f"，正文详情={upstream_p3_detail}"
    return passed, evidence


def structure_preservation_check(
    source: Presentation,
    candidate: Presentation,
) -> tuple[bool, str]:
    errors = []
    if source.slide_width != candidate.slide_width or source.slide_height != candidate.slide_height:
        errors.append("画布尺寸改变")
    for page, (source_slide, candidate_slide) in enumerate(
        zip(source.slides, candidate.slides), start=1
    ):
        if len(source_slide.shapes) != len(candidate_slide.shapes):
            errors.append(
                f"p{page}:对象数 {len(source_slide.shapes)}->{len(candidate_slide.shapes)}"
            )
            continue
        if background_signature(source_slide) != background_signature(candidate_slide):
            errors.append(f"p{page}:背景改变")
        for index, (source_shape, candidate_shape) in enumerate(
            zip(source_slide.shapes, candidate_slide.shapes)
        ):
            if source_shape.shape_type != candidate_shape.shape_type:
                errors.append(
                    f"p{page}/shape{index}:对象类型 {source_shape.shape_type}->{candidate_shape.shape_type}"
                )
                continue
            text_bearing = bool(getattr(source_shape, "has_text_frame", False))
            allowed_delta = 0.35 if text_bearing else 0.02
            delta = geometry_delta_inches(source_shape, candidate_shape)
            if delta > allowed_delta:
                errors.append(
                    f"p{page}/shape{index}:几何偏移 {delta:.3f}in>{allowed_delta:.2f}in"
                )
            if image_digest(source_shape) != image_digest(candidate_shape):
                errors.append(f"p{page}/shape{index}:图片二进制改变")
            if shape_style_signature(source_shape) != shape_style_signature(candidate_shape):
                errors.append(f"p{page}/shape{index}:填充、线条或旋转改变")
    passed = not errors
    return passed, "46 页对象、图片、背景和主体几何保持" if passed else "；".join(errors[:20])


def semantic_check(candidate: Presentation) -> tuple[bool, float, str]:
    passed_groups = 0
    total_groups = 0
    full_pages = 0
    missing = {}
    for page, groups in SEMANTIC_ANCHORS.items():
        content = normalized_text(slide_text(candidate.slides[page - 1]))
        page_missing = []
        for alternatives in groups:
            total_groups += 1
            normalized_alternatives = [normalized_text(value) for value in alternatives]
            if any(value in content for value in normalized_alternatives):
                passed_groups += 1
            else:
                page_missing.append("/".join(alternatives))
        if page_missing:
            missing[page] = page_missing
        else:
            full_pages += 1
    coverage = passed_groups / total_groups if total_groups else 0.0
    passed = coverage >= 0.85 and full_pages >= 38
    evidence = (
        f"语义锚点={passed_groups}/{total_groups}，完整页面={full_pages}/46"
        + (f"，缺失={dict(list(missing.items())[:10])}" if missing else "")
    )
    return passed, coverage, evidence


def run_style_match(source_run, candidate_run) -> bool:
    source_size = source_run.font.size.pt if source_run.font.size else None
    candidate_size = candidate_run.font.size.pt if candidate_run.font.size else None
    size_ok = (
        source_size is None
        and candidate_size is None
        or source_size is not None
        and candidate_size is not None
        and abs(source_size - candidate_size) <= 0.6
    )
    return (
        size_ok
        and source_run.font.bold == candidate_run.font.bold
        and source_run.font.italic == candidate_run.font.italic
        and source_run.font.name == candidate_run.font.name
        and color_signature(source_run.font.color)
        == color_signature(candidate_run.font.color)
    )


def text_style_check(source: Presentation, candidate: Presentation) -> tuple[bool, float, str]:
    matched = 0
    total = 0
    examples = []
    for page, (source_slide, candidate_slide) in enumerate(
        zip(source.slides, candidate.slides), start=1
    ):
        for index, source_shape in enumerate(source_slide.shapes):
            if (page, index) == (27, 16):
                continue
            source_runs = visible_runs(source_shape)
            if not source_runs:
                continue
            total += 1
            if index >= len(candidate_slide.shapes):
                examples.append(f"p{page}/shape{index}:对象缺失")
                continue
            candidate_runs = visible_runs(candidate_slide.shapes[index])
            if candidate_runs and run_style_match(source_runs[0], candidate_runs[0]):
                matched += 1
            else:
                examples.append(f"p{page}/shape{index}:首个文字 run 样式改变")
    ratio = matched / total if total else 0.0
    return ratio >= 0.95, ratio, f"首个可见文字 run 样式={matched}/{total}，示例={examples[:8]}"


def upstream_font_check(candidate: Presentation) -> tuple[bool, float, str]:
    checks = [
        ("封面标题", upstream_verifier.rule_m1_cover_title_font),
        ("封面副标题", upstream_verifier.rule_m1_cover_subtitle_font),
        ("章节页", upstream_verifier.rule_m5_title_slides_font),
        ("页脚", upstream_verifier.rule_m5_page_footer_font),
        ("目录", upstream_verifier.rule_m1_toc_font),
    ]
    passed = []
    failed = []
    for label, function in checks:
        penalty_hit, detail = function(candidate)
        if penalty_hit:
            failed.append(f"{label}:{detail}")
        else:
            passed.append(label)
    ratio = len(passed) / len(checks)
    return not failed, ratio, f"通过={passed}" + (f"，失败={failed}" if failed else "")


def render_check(path: Path) -> tuple[bool, str]:
    required_tools = ("libreoffice", "pdfinfo", "pdftotext")
    if not all(shutil.which(name) for name in required_tools):
        return False, "缺少 libreoffice、pdfinfo 或 pdftotext"
    with tempfile.TemporaryDirectory(prefix="cowork-evo-074-") as directory:
        converted = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", directory, str(path)],
            capture_output=True,
            text=True,
            timeout=240,
            check=False,
        )
        pdf = Path(directory) / f"{path.stem}.pdf"
        if converted.returncode != 0 or not pdf.is_file():
            return False, (converted.stderr or converted.stdout or "PDF 转换失败")[-1000:]
        info = subprocess.run(
            ["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=60, check=False
        )
        extracted = subprocess.run(
            ["pdftotext", "-layout", str(pdf), "-"],
            capture_output=True,
            text=True,
            timeout=90,
            check=False,
        )
        page_match = re.search(r"^Pages:\s+(\d+)", info.stdout, flags=re.M)
        pages = int(page_match.group(1)) if page_match else 0
        text = normalized_text(extracted.stdout)
    required = ["qinglan", "contents", "precision agriculture", "cotton", "platform architecture", "thank"]
    missing = [value for value in required if value not in text]
    return pages == 46 and not missing and len(text) > 10000, (
        f"渲染页数={pages}，提取字符={len(text)}，缺失={missing}"
    )


def professional_check(candidate: Presentation) -> tuple[bool, str]:
    errors = []
    tolerance = 50000
    for page, slide in enumerate(candidate.slides, start=1):
        if not any(slide_text(slide).strip() for _ in [0]):
            errors.append(f"p{page}:无可见文字")
        for shape in slide.shapes:
            if (
                shape.left < -tolerance
                or shape.top < -tolerance
                or shape.left + shape.width > candidate.slide_width + tolerance
                or shape.top + shape.height > candidate.slide_height + tolerance
            ):
                errors.append(f"p{page}:{shape.name} 越界")
            for run in visible_runs(shape):
                if run.font.size is not None and run.font.size.pt < 7.7:
                    errors.append(f"p{page}:{shape.name} 字号={run.font.size.pt:.2f}pt")
    return not errors, "对象均在页内，显式字号不低于 7.8pt" if not errors else "；".join(errors[:20])


def evaluate(artifact_dir: Path) -> dict:
    output = artifact_dir / OUTPUT_NAME
    results = []
    if not output.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        results.append(criterion("R009", True, "指定交付文件不存在"))
        return build_result(results, output)
    try:
        candidate = Presentation(output)
    except Exception as exc:
        results.append(criterion("R001", False, f"文件无法打开：{exc}"))
        results.append(criterion("R009", True, "交付 PPTX 已损坏"))
        return build_result(results, output)

    valid = len(candidate.slides) == 46
    results.append(criterion("R001", valid, f"页数={len(candidate.slides)}"))
    if not valid:
        results.append(criterion("R009", True, "页数不是 46"))
        return build_result(results, output)

    task_root = Path(__file__).resolve().parents[1]
    source = Presentation(task_root / "data" / "input_files" / SOURCE_NAME)

    english_ok, english_evidence = english_completion_check(source, candidate)
    results.append(criterion("R002", english_ok, english_evidence))

    structure_ok, structure_evidence = structure_preservation_check(source, candidate)
    results.append(criterion("R003", structure_ok, structure_evidence))

    semantic_ok, semantic_score, semantic_evidence = semantic_check(candidate)
    results.append(criterion("R004", semantic_ok, semantic_evidence, semantic_score))

    style_ok, style_score, style_evidence = text_style_check(source, candidate)
    results.append(criterion("R005", style_ok, style_evidence, style_score))

    font_ok, font_score, font_evidence = upstream_font_check(candidate)
    results.append(criterion("R006", font_ok, font_evidence, font_score))

    rendered, render_evidence = render_check(output)
    results.append(criterion("R007", rendered, render_evidence))

    professional, professional_evidence = professional_check(candidate)
    results.append(criterion("R008", professional, professional_evidence))

    results.append(
        criterion(
            "R009",
            not english_ok,
            "残留中文、空白替换或对象缺失" if not english_ok else "507 个源中文对象均已英文化",
        )
    )
    results.append(
        criterion(
            "R010",
            not structure_ok,
            "图片、背景、对象结构或主体版式改变" if not structure_ok else "图片、背景和主体版式保持",
        )
    )
    return build_result(results, output)


def build_result(results: list[dict], artifact_path: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "因交付不存在、损坏或页数错误，本判据未执行"))
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
    hurdles = [
        item["criterion_id"]
        for item in results
        if item["type"] == "hurdle" and not item["passed"]
    ]
    if hurdles:
        reward = 0.0

    digest = (
        hashlib.sha256(artifact_path.read_bytes()).hexdigest()
        if artifact_path.is_file()
        else hashlib.sha256(b"").hexdigest()
    )
    normalized_results = []
    for item in results:
        reported_score = (
            0.0 if item["passed"] else 1.0
        ) if item["type"] == "penalty" else item["score"]
        normalized_results.append(
            {
                "criterion_id": item["criterion_id"],
                "score": round(reported_score, 6),
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
        "criterion_results": normalized_results,
        "reward": round(reward, 6),
        "passed": reward >= PASS_THRESHOLD,
        "hurdle_failures": hurdles,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))

    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
