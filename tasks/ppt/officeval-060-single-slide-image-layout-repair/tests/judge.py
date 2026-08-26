#!/usr/bin/env python3
"""officeval_060 单页排版修复的独立 Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


OUTPUT_NAME = "PPT_slide4_relayout.pptx"
TASK_ID = "cowork-evo/officeval-060-single-slide-image-layout-repair"
RUBRIC_VERSION = "omegause-officeval-060-adaptation-v0.1"
JUDGE_VERSION = "officeval-060-adaptation-v0.1"
SLIDE4_XML = "ppt/slides/slide4.xml"
SPECS = {
    "R001": ("指定 PPTX 可打开，严格为 5 页 16:9", 4, "hurdle"),
    "R002": ("除第 4 页 XML 外，源 PPTX 其他包成员逐字节不变", 8, "hurdle"),
    "R003": ("第 4 页背景图全出血覆盖页面，四边留白不超过 0.03 cm", 6, "hurdle"),
    "R004": ("第 4 页背景图依原始像素比例上下对称裁剪到 16:9，没有拉伸", 5, "positive"),
    "R005": ("第 4 页标题、副标题、3 条说明和 4 个流程标签完整", 6, "hurdle"),
    "R006": ("第 4 页信息卡和全部文本在页面边界内", 4, "positive"),
    "R007": ("背景图仍为源文件中的同一张设备图，未替换或重编码", 4, "positive"),
    "R008": ("LibreOffice 可实际渲染为 5 页，第 4 页关键文本可读", 4, "positive"),
    "R009": ("若页数改变、非第4页成员被改写、背景未全出血或缺少两类内容，扣 15 分", -15, "penalty"),
}


def locate_source() -> Path | None:
    candidates = [Path("/app/input_files/PPT.pptx"), Path(__file__).resolve().parents[1] / "data/input_files/PPT.pptx"]
    return next((path for path in candidates if path.is_file()), None)


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {"criterion_id": identifier, "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic", "raw": {"description": description, "weight": weight, "type": kind, "passed": bool(passed)}}


def package_diff(source: Path, output: Path) -> tuple[bool, list[str]]:
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(output) as repaired:
        if set(original.namelist()) != set(repaired.namelist()):
            return False, ["package_member_set"]
        changed = [name for name in original.namelist() if original.read(name) != repaired.read(name)]
    return changed == [SLIDE4_XML], changed


def render(path: Path) -> tuple[bool, str]:
    with tempfile.TemporaryDirectory(prefix="officeval060-judge-") as temp_dir:
        completed = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, str(path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=150)
        pdf = Path(temp_dir) / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf.is_file(): return False, f"转 PDF 失败：{completed.stdout[-300:]}"
        info = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, text=True, timeout=30).stdout
        match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE); pages = int(match.group(1)) if match else 0
        txt = Path(temp_dir) / "out.txt"; subprocess.run(["pdftotext", "-f", "4", "-l", "4", str(pdf), str(txt)], check=True, timeout=30)
        rendered = re.sub(r"\s+", "", txt.read_text(errors="replace")); keys = ["多级膜分离与清水回用链路", "预处理", "膜分离", "压差判读", "回用缓冲"]
        ok = pages == 5 and all(re.sub(r"\s+", "", key) in rendered for key in keys)
        return ok, f"页数={pages}，第4页关键文本可读={ok}"


def evaluate(path: Path) -> dict:
    criteria = []; source = locate_source()
    try: prs = Presentation(path)
    except Exception as exc:
        for identifier in SPECS: criteria.append(criterion(identifier, False, f"打开失败：{exc}"))
        return make_result(path, criteria)
    size_ok = abs(prs.slide_width / 360000 - 33.8667) <= .03 and abs(prs.slide_height / 360000 - 19.05) <= .03
    criteria.append(criterion("R001", len(prs.slides) == 5 and size_ok, f"页数={len(prs.slides)}，页面={prs.slide_width/360000:.3f}×{prs.slide_height/360000:.2f}cm"))
    package_ok, changed = package_diff(source, path) if source else (False, ["source_missing"])
    criteria.append(criterion("R002", package_ok, f"发生字节变化的包成员={changed}"))
    slide4 = prs.slides[3] if len(prs.slides) >= 4 else None
    pictures = [item for item in slide4.shapes if item.shape_type == MSO_SHAPE_TYPE.PICTURE] if slide4 else []
    picture = pictures[0] if len(pictures) == 1 else None
    full_bleed = bool(picture and abs(picture.left) <= 1000 and abs(picture.top) <= 1000 and abs(picture.width - prs.slide_width) <= 1000 and abs(picture.height - prs.slide_height) <= 1000)
    criteria.append(criterion("R003", full_bleed, f"图片数={len(pictures)}，全出血={full_bleed}"))
    ratio_ok = False; crop_evidence = ""
    if picture:
        with Image.open(BytesIO(picture.image.blob)) as image:
            visible_w = image.width * (1 - picture.crop_left - picture.crop_right)
            visible_h = image.height * (1 - picture.crop_top - picture.crop_bottom)
            visible_ratio = visible_w / visible_h
            ratio_ok = abs(visible_ratio - 16 / 9) <= .003 and abs(picture.crop_top - picture.crop_bottom) <= .0001 and picture.crop_left == 0 and picture.crop_right == 0
            crop_evidence = f"像素={image.width}×{image.height}，裁剪={picture.crop_left:.4f}/{picture.crop_top:.4f}/{picture.crop_right:.4f}/{picture.crop_bottom:.4f}，可见比={visible_ratio:.4f}"
    criteria.append(criterion("R004", ratio_ok, crop_evidence or "无背景图"))
    content = "\n".join(item.text for item in slide4.shapes if getattr(item, "has_text_frame", False)) if slide4 else ""
    normalized = re.sub(r"\s+", "", content)
    required = ["多级膜分离与清水回用链路", "预处理、膜组件与透明水箱构成连续式中试平台", "左侧高浊度水进入预处理罐", "中段膜堆完成深度分离", "右侧清水箱作为回用缓冲", "预处理", "膜分离", "压差判读", "回用缓冲", "工艺说明"]
    missing = [value for value in required if re.sub(r"\s+", "", value) not in normalized]
    criteria.append(criterion("R005", not missing, f"缺失={missing}"))
    inside = slide4 and all(item.left >= -1000 and item.top >= -1000 and item.left + item.width <= prs.slide_width + 1000 and item.top + item.height <= prs.slide_height + 1000 for item in slide4.shapes)
    criteria.append(criterion("R006", bool(inside), f"第4页全部对象在页内={bool(inside)}"))
    image_same = False
    if picture and source:
        source_prs = Presentation(source); source_pictures = [item for item in source_prs.slides[3].shapes if item.shape_type == MSO_SHAPE_TYPE.PICTURE]
        image_same = len(source_pictures) == 1 and hashlib.sha256(picture.image.blob).digest() == hashlib.sha256(source_pictures[0].image.blob).digest()
    criteria.append(criterion("R007", image_same, f"背景图媒体字节不变={image_same}"))
    render_ok, render_evidence = render(path); criteria.append(criterion("R008", render_ok, render_evidence))
    severe = len(prs.slides) != 5 or not package_ok or not full_bleed or len(missing) >= 2
    criteria.append(criterion("R009", not severe, f"严重错误={severe}"))
    return make_result(path, criteria)


def make_result(path: Path, criteria: list[dict]) -> dict:
    hurdles = [item["criterion_id"] for item in criteria if item["raw"]["type"] == "hurdle" and not item["raw"]["passed"]]
    total = sum(item["raw"]["weight"] for item in criteria if item["raw"]["weight"] > 0); earned = sum(item["raw"]["weight"] for item in criteria if item["raw"]["weight"] > 0 and item["raw"]["passed"]); penalties = sum(abs(item["raw"]["weight"]) for item in criteria if item["raw"]["weight"] < 0 and not item["raw"]["passed"])
    reward = 0.0 if hurdles else max(0.0, min(1.0, (earned - penalties) / total))
    return {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [hashlib.sha256(path.read_bytes()).hexdigest()] if path.is_file() else [], "criterion_results": criteria, "reward": reward, "passed": reward >= .8 and not hurdles, "hurdle_failures": hurdles}


def main() -> None:
    parser = argparse.ArgumentParser(); parser.add_argument("--artifact-dir", type=Path, required=True); parser.add_argument("--result", type=Path, required=True); parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args(); output = args.artifact_dir / OUTPUT_NAME
    evaluation = evaluate(output) if output.is_file() else {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [], "criterion_results": [], "reward": 0.0, "passed": False, "hurdle_failures": ["missing_output"]}
    args.result.parent.mkdir(parents=True, exist_ok=True); args.reward_file.parent.mkdir(parents=True, exist_ok=True); args.result.write_text(json.dumps(evaluation, ensure_ascii=False, indent=2), encoding="utf-8"); args.reward_file.write_text(f"{evaluation['reward']:.6f}\n", encoding="utf-8"); print(json.dumps(evaluation, ensure_ascii=False, indent=2))


if __name__ == "__main__": main()
