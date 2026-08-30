#!/usr/bin/env python3
"""PPTArena Case 16 的任务专属确定性 Judge。"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import re
import subprocess
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree
from PIL import Image, ImageChops
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


OUTPUT_NAME = "OrderPhotosGroundTruth.pptx"
INPUT_NAME = "OrderPhotosTest.pptx"
INPUT_SHA256 = "c17f8b37992c462006bf62a151b3f346df2dade7253f97027f69d715ab7f657c"
TASK_ID = "cowork-evo/pptarena-016-aligned-photo-triptych"
RUBRIC_VERSION = "pptarena-case-016-adaptation-v0.1"
JUDGE_VERSION = "pptarena-case-016-deterministic-v0.1"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PER_POINT = 12_700
SOURCE_SHAPE_COUNTS = [2, 12]
SOURCE_MEMBER_COUNT = 43
TARGETS = {
    "8d1cb44a11c4fd45c997eec889d9b40e12225fba0978fdebbbc41bab173327d3": {
        "semantic": "WATERFALL",
        "shape_name": "Picture 26",
        "left_pt": 75.0,
    },
    "3667007504e1352212424e7e01b4b8d4f19a851831ddbb9c3861717b55699ed9": {
        "semantic": "STAIRCASE",
        "shape_name": "Picture 4",
        "left_pt": 354.39,
    },
    "f9724655f4ae12269ee673944af1e1329dc103e32a2917d9c7d9b2c54b1bf47b": {
        "semantic": "CHESS",
        "shape_name": "Picture 18",
        "left_pt": 632.83,
    },
}
SUSPICIOUS_MARKERS = [
    "solution/golden",
    "/root/",
    "judge-result",
    "artifact_dir",
    "pptarena-case-016",
]
SPECS = {
    "R001": ("指定 PPTX 可打开、页面尺寸不变且总页数为 2", 6, "hurdle"),
    "R002": ("三张原图按瀑布、楼梯、国际象从左到右与标签正确对应", 18, "hurdle"),
    "R003": ("三图无旋转、等大近方形、顶边对齐且不遮挡标签", 18, "hurdle"),
    "R004": ("第 2 页为显式的白色或主题背景白纯色填充", 8, "hurdle"),
    "R005": ("第 2 页除目标图与背景外的原对象、样式、顺序和备注保持", 14, "hurdle"),
    "R006": ("封面、主题、母版与版式保持，无无关换肤", 12, "hurdle"),
    "R007": ("完整渲染 2 页，封面和非目标区域视觉保持", 10, "hurdle"),
    "R008": ("保留 16 个原生对象、4 个原媒体和三图原裁剪，无压平或额外素材", 6, "hurdle"),
    "R009": ("输出目录干净且无答案路径或未授权文字", 5, "positive"),
    "R010": ("规范化合法改动后内容 XML 和关键包部件字节保持", 3, "positive"),
    "R011": ("图片匹配、排布、可读性、背景或原内容保持发生严重错误", -20, "penalty"),
}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def locate_source() -> Path | None:
    candidates = [
        Path("/app/input_files") / INPUT_NAME,
        Path(__file__).resolve().parents[1] / "data/input_files" / INPUT_NAME,
    ]
    return next((path for path in candidates if path.is_file()), None)


def package_ok(path: Path) -> tuple[bool, str]:
    try:
        with zipfile.ZipFile(path) as archive:
            broken = archive.testzip()
            names = {name for name in archive.namelist() if not name.endswith("/")}
            required = {
                "[Content_Types].xml",
                "ppt/presentation.xml",
                "ppt/slides/slide1.xml",
                "ppt/slides/slide2.xml",
                "ppt/theme/theme1.xml",
            }
            missing = sorted(required - names)
            macros = sorted(name for name in names if name.lower().endswith("vbaproject.bin"))
        return broken is None and not missing and not macros, (
            f"损坏成员={broken}，缺失={missing}，宏={macros}"
        )
    except Exception as exc:
        return False, f"ZIP 检查失败：{exc}"


def picture_hash(shape) -> str:
    try:
        return hashlib.sha256(shape.image.blob).hexdigest()
    except Exception:
        return "unreadable"


def picture_crop(shape) -> tuple[float, float, float, float]:
    return tuple(
        round(float(value), 6)
        for value in (shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom)
    )


def shape_element_hash(shape) -> str:
    return hashlib.sha256(etree.tostring(shape.element, method="c14n")).hexdigest()


def shape_inventory(slide, *, exclude_target_pictures: bool) -> list[tuple]:
    values = []
    for index, shape in enumerate(slide.shapes):
        digest = picture_hash(shape) if shape.shape_type == MSO_SHAPE_TYPE.PICTURE else None
        if exclude_target_pictures and digest in TARGETS:
            continue
        values.append((index, int(shape.shape_id), shape.name, str(shape.shape_type), shape_element_hash(shape)))
    return values


def notes_signature(presentation: Presentation) -> tuple[str, ...]:
    values = []
    for slide in presentation.slides:
        try:
            values.append(slide.notes_slide.notes_text_frame.text)
        except Exception:
            values.append("")
    return tuple(values)


def media_hashes(path: Path) -> Counter[str]:
    with zipfile.ZipFile(path) as archive:
        return Counter(
            hashlib.sha256(archive.read(name)).hexdigest()
            for name in archive.namelist()
            if name.startswith("ppt/media/") and not name.endswith("/")
        )


def member_names(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as archive:
        return {name for name in archive.namelist() if not name.endswith("/")}


def xml_subtree_hash(path: Path, part: str, xpath: str, namespaces: dict[str, str]) -> str:
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read(part))
    nodes = root.xpath(xpath, namespaces=namespaces)
    if len(nodes) != 1:
        return f"missing:{len(nodes)}"
    return hashlib.sha256(etree.tostring(nodes[0], method="c14n")).hexdigest()


def xml_family_hashes(path: Path, prefixes: tuple[str, ...]) -> dict[str, str]:
    values = {}
    with zipfile.ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            if name.endswith(".xml") and name.startswith(prefixes):
                root = etree.fromstring(archive.read(name))
                values[name] = hashlib.sha256(etree.tostring(root, method="c14n")).hexdigest()
    return values


def target_inventory(presentation: Presentation) -> tuple[dict[str, dict], list[str]]:
    values: dict[str, dict] = {}
    all_picture_hashes = []
    if len(presentation.slides) < 2:
        return values, all_picture_hashes
    for z_index, shape in enumerate(presentation.slides[1].shapes):
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        digest = picture_hash(shape)
        all_picture_hashes.append(digest)
        if digest not in TARGETS:
            continue
        if digest in values:
            values[digest]["duplicate"] = True
            continue
        rotation = float(shape.rotation or 0) % 360
        rotation = min(rotation, abs(360 - rotation))
        values[digest] = {
            "semantic": TARGETS[digest]["semantic"],
            "shape_id": int(shape.shape_id),
            "name": shape.name,
            "z_index": z_index,
            "left_pt": shape.left / EMU_PER_POINT,
            "top_pt": shape.top / EMU_PER_POINT,
            "width_pt": shape.width / EMU_PER_POINT,
            "height_pt": shape.height / EMU_PER_POINT,
            "right_pt": (shape.left + shape.width) / EMU_PER_POINT,
            "bottom_pt": (shape.top + shape.height) / EMU_PER_POINT,
            "rotation": rotation,
            "crop": picture_crop(shape),
            "duplicate": False,
            "shape": shape,
        }
    return values, all_picture_hashes


def label_inventory(slide) -> dict[str, dict]:
    values = {}
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = shape.text.strip()
        if text in {"WATERFALL", "STAIRCASE", "CHESS"}:
            values[text] = {
                "shape_id": int(shape.shape_id),
                "left_pt": shape.left / EMU_PER_POINT,
                "top_pt": shape.top / EMU_PER_POINT,
                "width_pt": shape.width / EMU_PER_POINT,
                "height_pt": shape.height / EMU_PER_POINT,
            }
    return values


def panel_geometry(slide) -> tuple[float, float, float, float] | None:
    for shape in slide.shapes:
        if int(shape.shape_id) == 28 and shape.name == "Rectangle 27":
            return (
                shape.left / EMU_PER_POINT,
                shape.top / EMU_PER_POINT,
                (shape.left + shape.width) / EMU_PER_POINT,
                (shape.top + shape.height) / EMU_PER_POINT,
            )
    return None


def background_info(path: Path) -> dict:
    try:
        with zipfile.ZipFile(path) as archive:
            root = etree.fromstring(archive.read("ppt/slides/slide2.xml"))
        namespaces = {"p": P_NS, "a": A_NS}
        fills = root.xpath("./p:cSld/p:bg/p:bgPr/a:solidFill", namespaces=namespaces)
        if len(fills) != 1:
            return {"passed": False, "reason": f"显式 solidFill 数={len(fills)}"}
        children = [child for child in fills[0] if isinstance(child.tag, str)]
        if len(children) != 1:
            return {"passed": False, "reason": f"solidFill 颜色节点数={len(children)}"}
        child = children[0]
        local = etree.QName(child).localname
        value = child.get("val", "").upper()
        passed = False
        if local == "schemeClr":
            passed = value.lower() in {"bg1", "lt1"}
        elif local == "srgbClr" and re.fullmatch(r"[0-9A-F]{6}", value):
            channels = [int(value[index : index + 2], 16) for index in (0, 2, 4)]
            passed = min(channels) >= 245 and max(channels) - min(channels) <= 8
        return {"passed": passed, "color_type": local, "value": value}
    except Exception as exc:
        return {"passed": False, "reason": str(exc)}


def normalized_content_hash(path: Path, source: Path) -> tuple[str, str]:
    namespaces = {"p": P_NS, "a": A_NS}
    with zipfile.ZipFile(source) as source_archive, zipfile.ZipFile(path) as candidate_archive:
        source_parts = []
        candidate_parts = []
        for number in (1, 2):
            source_root = etree.fromstring(source_archive.read(f"ppt/slides/slide{number}.xml"))
            candidate_root = etree.fromstring(candidate_archive.read(f"ppt/slides/slide{number}.xml"))
            if number == 2:
                for root in (source_root, candidate_root):
                    for background in root.xpath("./p:cSld/p:bg", namespaces=namespaces):
                        background.getparent().remove(background)
                for target in TARGETS.values():
                    name = target["shape_name"]
                    source_nodes = source_root.xpath(
                        f"./p:cSld/p:spTree/p:pic[p:nvPicPr/p:cNvPr/@name='{name}']/p:spPr/a:xfrm",
                        namespaces=namespaces,
                    )
                    candidate_nodes = candidate_root.xpath(
                        f"./p:cSld/p:spTree/p:pic[p:nvPicPr/p:cNvPr/@name='{name}']/p:spPr/a:xfrm",
                        namespaces=namespaces,
                    )
                    if len(source_nodes) != 1 or len(candidate_nodes) != 1:
                        raise ValueError(f"无法规范化图片 {name}")
                    candidate_nodes[0].getparent().replace(candidate_nodes[0], copy.deepcopy(source_nodes[0]))
            source_parts.append(etree.tostring(source_root.find(qn(P_NS, "cSld")), method="c14n"))
            candidate_parts.append(etree.tostring(candidate_root.find(qn(P_NS, "cSld")), method="c14n"))
    return (
        hashlib.sha256(b"\n".join(source_parts)).hexdigest(),
        hashlib.sha256(b"\n".join(candidate_parts)).hexdigest(),
    )


def critical_hashes(path: Path) -> dict[str, str]:
    excluded = {
        "ppt/slides/slide2.xml",
        "docProps/app.xml",
        "docProps/core.xml",
        "docProps/thumbnail.jpeg",
    }
    with zipfile.ZipFile(path) as archive:
        return {
            name: hashlib.sha256(archive.read(name)).hexdigest()
            for name in sorted(archive.namelist())
            if not name.endswith("/") and name not in excluded
        }


def render_presentation(path: Path, outdir: Path) -> tuple[list[Path], int, str]:
    outdir.mkdir(parents=True, exist_ok=True)
    profile = outdir / "lo-profile"
    profile.mkdir()
    completed = subprocess.run(
        [
            "soffice",
            f"-env:UserInstallation=file://{profile}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(path),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        timeout=240,
        check=False,
    )
    pdf = outdir / f"{path.stem}.pdf"
    if completed.returncode != 0 or not pdf.is_file():
        return [], 0, f"转 PDF 失败：{completed.stdout[-500:]}"
    info = subprocess.run(
        ["pdfinfo", str(pdf)],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        timeout=30,
        check=False,
    ).stdout
    match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
    pages = int(match.group(1)) if match else 0
    prefix = outdir / "slide"
    raster = subprocess.run(
        ["pdftoppm", "-jpeg", "-r", "120", str(pdf), str(prefix)],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        timeout=240,
        check=False,
    )
    images = sorted(outdir.glob("slide-*.jpg")) if raster.returncode == 0 else []
    return images, pages, (
        f"页数={pages}，渲染图={len(images)}，soffice={completed.returncode}，pdftoppm={raster.returncode}"
    )


def rotated_bbox(shape) -> tuple[float, float, float, float]:
    left = float(shape.left)
    top = float(shape.top)
    width = float(shape.width)
    height = float(shape.height)
    angle = math.radians(float(shape.rotation or 0))
    center_x = left + width / 2
    center_y = top + height / 2
    corners = []
    for x, y in ((left, top), (left + width, top), (left + width, top + height), (left, top + height)):
        dx = x - center_x
        dy = y - center_y
        corners.append(
            (
                center_x + dx * math.cos(angle) - dy * math.sin(angle),
                center_y + dx * math.sin(angle) + dy * math.cos(angle),
            )
        )
    return (
        min(value[0] for value in corners),
        min(value[1] for value in corners),
        max(value[0] for value in corners),
        max(value[1] for value in corners),
    )


def masked_similarity(
    source_image: Image.Image,
    candidate_image: Image.Image,
    source_shapes: list,
    candidate_shapes: list,
    slide_size: tuple[int, int],
) -> tuple[float, dict]:
    source_rgb = source_image.convert("RGB")
    candidate_rgb = candidate_image.convert("RGB")
    if source_rgb.size != candidate_rgb.size:
        return 0.0, {"reason": "渲染尺寸不一致"}
    mask = Image.new("L", source_rgb.size, 255)
    padding = max(5, round(min(source_rgb.size) * 0.008))
    rectangles = []
    for shape in [*source_shapes, *candidate_shapes]:
        left, top, right, bottom = rotated_bbox(shape)
        box = (
            round(left * source_rgb.width / slide_size[0]) - padding,
            round(top * source_rgb.height / slide_size[1]) - padding,
            round(right * source_rgb.width / slide_size[0]) + padding,
            round(bottom * source_rgb.height / slide_size[1]) + padding,
        )
        clipped = (
            max(0, box[0]),
            max(0, box[1]),
            min(source_rgb.width, box[2]),
            min(source_rgb.height, box[3]),
        )
        mask.paste(0, clipped)
        rectangles.append(clipped)
    difference = ImageChops.difference(source_rgb, candidate_rgb)
    pixels = mask.histogram()[255]
    histogram = difference.histogram(mask)
    squared = sum(count * ((index % 256) ** 2) for index, count in enumerate(histogram))
    rms = math.sqrt(squared / (pixels * 3)) if pixels else 255.0
    return max(0.0, 1.0 - rms / 255.0), {
        "masked_rectangles": rectangles,
        "evaluated_pixels": pixels,
        "rms": round(rms, 6),
    }


def whole_image_similarity(source_image: Image.Image, candidate_image: Image.Image) -> tuple[float, dict]:
    source_rgb = source_image.convert("RGB")
    candidate_rgb = candidate_image.convert("RGB")
    if source_rgb.size != candidate_rgb.size:
        return 0.0, {"reason": "渲染尺寸不一致"}
    difference = ImageChops.difference(source_rgb, candidate_rgb)
    histogram = difference.histogram()
    pixels = source_rgb.width * source_rgb.height
    squared = sum(count * ((index % 256) ** 2) for index, count in enumerate(histogram))
    rms = math.sqrt(squared / (pixels * 3)) if pixels else 255.0
    return max(0.0, 1.0 - rms / 255.0), {"rms": round(rms, 6), "pixels": pixels}


def result_item(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {
        "criterion_id": identifier,
        "score": 1.0 if passed else 0.0,
        "evidence": evidence,
        "method": "deterministic",
        "raw": {
            "description": description,
            "weight": weight,
            "type": kind,
            "passed": bool(passed),
        },
    }


def make_result(path: Path, criteria: list[dict]) -> dict:
    hurdles = [
        item["criterion_id"]
        for item in criteria
        if item["raw"]["type"] == "hurdle" and not item["raw"]["passed"]
    ]
    total = sum(item["raw"]["weight"] for item in criteria if item["raw"]["weight"] > 0)
    earned = sum(
        item["raw"]["weight"]
        for item in criteria
        if item["raw"]["weight"] > 0 and item["raw"]["passed"]
    )
    penalty = sum(
        abs(item["raw"]["weight"])
        for item in criteria
        if item["raw"]["weight"] < 0 and not item["raw"]["passed"]
    )
    reward = 0.0 if hurdles else max(0.0, min(1.0, (earned - penalty) / total))
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [sha256(path)] if path.is_file() else [],
        "criterion_results": criteria,
        "reward": reward,
        "passed": reward >= 0.8 and not hurdles,
        "hurdle_failures": hurdles,
    }


def evaluate(path: Path, artifact_dir: Path) -> dict:
    source_path = locate_source()
    if source_path is None or sha256(source_path) != INPUT_SHA256:
        return make_result(
            path,
            [result_item(identifier, False, "冻结输入缺失或 SHA256 不匹配") for identifier in SPECS],
        )
    try:
        source = Presentation(source_path)
        candidate = Presentation(path)
    except Exception as exc:
        return make_result(
            path,
            [result_item(identifier, False, f"PPTX 打开失败：{exc}") for identifier in SPECS],
        )

    criteria: list[dict] = []
    package_passed, package_evidence = package_ok(path)
    size_passed = (
        len(source.slides) == 2
        and len(candidate.slides) == 2
        and candidate.slide_width == source.slide_width
        and candidate.slide_height == source.slide_height
    )
    criteria.append(
        result_item(
            "R001",
            package_passed and size_passed,
            f"{package_evidence}，输入页数={len(source.slides)}，候选页数={len(candidate.slides)}，"
            f"页面={candidate.slide_width / EMU_PER_POINT:.2f}×{candidate.slide_height / EMU_PER_POINT:.2f}pt",
        )
    )

    source_targets, source_picture_hashes = target_inventory(source)
    candidate_targets, candidate_picture_hashes = target_inventory(candidate)
    expected_hashes = set(TARGETS)
    mapping_passed = (
        set(source_targets) == expected_hashes
        and set(candidate_targets) == expected_hashes
        and len(candidate_picture_hashes) == 3
        and not any(item["duplicate"] for item in candidate_targets.values())
        and all(
            candidate_targets[digest]["name"] == TARGETS[digest]["shape_name"]
            and abs(candidate_targets[digest]["left_pt"] - TARGETS[digest]["left_pt"]) <= 6.0
            for digest in expected_hashes
        )
        and [
            item["semantic"]
            for item in sorted(candidate_targets.values(), key=lambda value: value["left_pt"])
        ]
        == ["WATERFALL", "STAIRCASE", "CHESS"]
    )
    target_evidence = {
        item["semantic"]: {
            key: round(value, 4) if isinstance(value, float) else value
            for key, value in item.items()
            if key not in {"shape", "crop"}
        }
        for item in candidate_targets.values()
    }
    criteria.append(
        result_item(
            "R002",
            mapping_passed,
            f"第2页图片数={len(candidate_picture_hashes)}，目标={json.dumps(target_evidence, ensure_ascii=False)}",
        )
    )

    labels = label_inventory(candidate.slides[1]) if len(candidate.slides) == 2 else {}
    panel = panel_geometry(candidate.slides[1]) if len(candidate.slides) == 2 else None
    geometry_checks = []
    if set(candidate_targets) == expected_hashes and set(labels) == {"WATERFALL", "STAIRCASE", "CHESS"} and panel:
        for digest, target in TARGETS.items():
            item = candidate_targets[digest]
            label = labels[target["semantic"]]
            gap = label["top_pt"] - item["bottom_pt"]
            geometry_checks.append(
                {
                    "semantic": target["semantic"],
                    "square": abs(item["width_pt"] - item["height_pt"]) <= 1.0,
                    "size": 224.5 <= item["width_pt"] <= 235.5 and 224.5 <= item["height_pt"] <= 235.5,
                    "top": 133.5 <= item["top_pt"] <= 136.5,
                    "rotation": item["rotation"] <= 0.2,
                    "label_gap_pt": round(gap, 4),
                    "label_gap_ok": 1.0 <= gap <= 15.0,
                    "inside_panel": (
                        item["left_pt"] >= panel[0]
                        and item["top_pt"] >= panel[1]
                        and item["right_pt"] <= panel[2]
                        and item["bottom_pt"] <= panel[3]
                    ),
                }
            )
        sizes = [item["width_pt"] for item in candidate_targets.values()]
        tops = [item["top_pt"] for item in candidate_targets.values()]
        ordered = sorted(candidate_targets.values(), key=lambda value: value["left_pt"])
        horizontal_gaps = [ordered[index + 1]["left_pt"] - ordered[index]["right_pt"] for index in range(2)]
        geometry_passed = (
            all(all(value for key, value in check.items() if key in {"square", "size", "top", "rotation", "label_gap_ok", "inside_panel"}) for check in geometry_checks)
            and max(sizes) - min(sizes) <= 3.0
            and max(tops) - min(tops) <= 1.0
            and min(horizontal_gaps) >= 30.0
        )
    else:
        sizes = []
        tops = []
        horizontal_gaps = []
        geometry_passed = False
    criteria.append(
        result_item(
            "R003",
            geometry_passed,
            f"逐图={json.dumps(geometry_checks, ensure_ascii=False)}，尺寸差="
            f"{round(max(sizes) - min(sizes), 4) if sizes else None}pt，顶边差="
            f"{round(max(tops) - min(tops), 4) if tops else None}pt，水平间隙="
            f"{[round(value, 4) for value in horizontal_gaps]}",
        )
    )

    background = background_info(path)
    background_type_passed = len(candidate.slides) == 2 and "SOLID" in str(candidate.slides[1].background.fill.type).upper()
    criteria.append(
        result_item(
            "R004",
            background_type_passed and bool(background.get("passed")),
            f"python-pptx 背景={candidate.slides[1].background.fill.type if len(candidate.slides) == 2 else None}，"
            f"OOXML={json.dumps(background, ensure_ascii=False)}",
        )
    )

    slide2_preserved = (
        len(candidate.slides) == 2
        and shape_inventory(source.slides[1], exclude_target_pictures=True)
        == shape_inventory(candidate.slides[1], exclude_target_pictures=True)
        and notes_signature(source)[1:] == notes_signature(candidate)[1:]
    )
    criteria.append(
        result_item(
            "R005",
            slide2_preserved,
            f"非目标对象签名保持={slide2_preserved}，"
            f"数量={len(shape_inventory(source.slides[1], exclude_target_pictures=True))}→"
            f"{len(shape_inventory(candidate.slides[1], exclude_target_pictures=True)) if len(candidate.slides) == 2 else None}，"
            f"第2页备注保持={notes_signature(source)[1:] == notes_signature(candidate)[1:]}",
        )
    )

    theme_source = xml_subtree_hash(source_path, "ppt/theme/theme1.xml", "./a:themeElements", {"a": A_NS})
    theme_candidate = xml_subtree_hash(path, "ppt/theme/theme1.xml", "./a:themeElements", {"a": A_NS})
    source_layouts = xml_family_hashes(source_path, ("ppt/slideMasters/", "ppt/slideLayouts/"))
    candidate_layouts = xml_family_hashes(path, ("ppt/slideMasters/", "ppt/slideLayouts/"))
    slide1_preserved = (
        len(candidate.slides) == 2
        and shape_inventory(source.slides[0], exclude_target_pictures=False)
        == shape_inventory(candidate.slides[0], exclude_target_pictures=False)
        and str(source.slides[0].background.fill.type) == str(candidate.slides[0].background.fill.type)
        and notes_signature(source)[:1] == notes_signature(candidate)[:1]
        and theme_source == theme_candidate
        and source_layouts == candidate_layouts
    )
    criteria.append(
        result_item(
            "R006",
            slide1_preserved,
            f"封面对象保持="
            f"{len(candidate.slides) == 2 and shape_inventory(source.slides[0], exclude_target_pictures=False) == shape_inventory(candidate.slides[0], exclude_target_pictures=False)}，"
            f"背景={source.slides[0].background.fill.type}→{candidate.slides[0].background.fill.type if len(candidate.slides) == 2 else None}，"
            f"主题保持={theme_source == theme_candidate}，母版/版式保持={source_layouts == candidate_layouts}",
        )
    )

    try:
        source_media = media_hashes(source_path)
        candidate_media = media_hashes(path)
        source_members = member_names(source_path)
        candidate_members = member_names(path)
    except Exception:
        source_media = candidate_media = Counter()
        source_members = candidate_members = set()
    crops_equal = set(candidate_targets) == expected_hashes and all(
        candidate_targets[digest]["crop"] == source_targets[digest]["crop"]
        for digest in expected_hashes
    )
    editable_passed = (
        package_passed
        and len(candidate.slides) == 2
        and [len(slide.shapes) for slide in candidate.slides] == SOURCE_SHAPE_COUNTS
        and source_media == candidate_media
        and sum(candidate_media.values()) == 4
        and source_members == candidate_members
        and len(candidate_members) == SOURCE_MEMBER_COUNT
        and crops_equal
    )
    criteria.append(
        result_item(
            "R008",
            editable_passed,
            f"对象数={[len(slide.shapes) for slide in candidate.slides]}，媒体哈希保持={source_media == candidate_media}，"
            f"媒体数={sum(source_media.values())}→{sum(candidate_media.values())}，包成员="
            f"{len(source_members)}→{len(candidate_members)}，裁剪保持={crops_equal}",
        )
    )

    pre_render_ready = all(
        item["raw"]["passed"]
        for item in criteria
        if item["criterion_id"] in {"R001", "R002", "R003", "R004", "R005", "R006", "R008"}
    )
    render_passed = False
    render_evidence = {"note": "前置检查失败，未执行渲染"}
    if pre_render_ready:
        with tempfile.TemporaryDirectory(prefix="pptarena016-judge-") as temporary:
            root = Path(temporary)
            source_images, source_pages, source_note = render_presentation(source_path, root / "source")
            candidate_images, candidate_pages, candidate_note = render_presentation(path, root / "candidate")
            if len(source_images) == 2 and len(candidate_images) == 2:
                with Image.open(source_images[0]) as source_cover, Image.open(candidate_images[0]) as candidate_cover:
                    cover_similarity, cover_detail = whole_image_similarity(source_cover, candidate_cover)
                with Image.open(source_images[1]) as source_page, Image.open(candidate_images[1]) as candidate_page:
                    slide2_similarity, slide2_detail = masked_similarity(
                        source_page,
                        candidate_page,
                        [item["shape"] for item in source_targets.values()],
                        [item["shape"] for item in candidate_targets.values()],
                        (source.slide_width, source.slide_height),
                    )
                render_passed = (
                    source_pages == 2
                    and candidate_pages == 2
                    and cover_similarity >= 0.9995
                    and slide2_similarity >= 0.9985
                )
                render_evidence = {
                    "source": source_note,
                    "candidate": candidate_note,
                    "cover_similarity": round(cover_similarity, 6),
                    "cover_detail": cover_detail,
                    "slide2_masked_similarity": round(slide2_similarity, 6),
                    "slide2_detail": slide2_detail,
                }
            else:
                render_evidence = {"source": source_note, "candidate": candidate_note}
    criteria.append(result_item("R007", render_passed, json.dumps(render_evidence, ensure_ascii=False)))

    artifact_files = sorted(
        str(item.relative_to(artifact_dir))
        for item in artifact_dir.rglob("*")
        if item.is_file()
    )
    all_text = "\n".join(
        shape.text
        for slide in candidate.slides
        for shape in slide.shapes
        if getattr(shape, "has_text_frame", False)
    )
    markers = [marker for marker in SUSPICIOUS_MARKERS if marker.lower() in all_text.lower()]
    hygiene_passed = artifact_files == [OUTPUT_NAME] and not markers
    criteria.append(
        result_item(
            "R009",
            hygiene_passed,
            f"输出文件={artifact_files}，可疑标记={markers}",
        )
    )

    try:
        source_normalized, candidate_normalized = normalized_content_hash(path, source_path)
        normalized_equal = source_normalized == candidate_normalized
        critical_equal = critical_hashes(path) == critical_hashes(source_path)
    except Exception as exc:
        normalized_equal = critical_equal = False
        source_normalized = candidate_normalized = f"error:{exc}"
    criteria.append(
        result_item(
            "R010",
            normalized_equal and critical_equal,
            f"规范化内容={source_normalized}→{candidate_normalized}，关键包部件保持={critical_equal}",
        )
    )

    severe = any(
        not item["raw"]["passed"]
        for item in criteria
        if item["criterion_id"] in {"R001", "R002", "R003", "R004", "R005", "R006", "R007", "R008"}
    ) or bool(markers)
    criteria.append(result_item("R011", not severe, f"严重错误={severe}"))
    order = {identifier: index for index, identifier in enumerate(SPECS)}
    criteria.sort(key=lambda item: order[item["criterion_id"]])
    return make_result(path, criteria)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    output = args.artifact_dir / OUTPUT_NAME
    if output.is_file():
        evaluation = evaluate(output, args.artifact_dir)
    else:
        evaluation = {
            "task_id": TASK_ID,
            "rubric_version": RUBRIC_VERSION,
            "judge_version": JUDGE_VERSION,
            "artifact_sha256": [],
            "criterion_results": [],
            "reward": 0.0,
            "passed": False,
            "hurdle_failures": ["missing_output"],
        }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(evaluation, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.write_text(f"{evaluation['reward']:.6f}\n", encoding="utf-8")
    print(json.dumps(evaluation, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
