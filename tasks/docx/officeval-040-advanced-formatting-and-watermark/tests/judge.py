#!/usr/bin/env python3
"""officeval_040 的独立 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from PIL import Image
from docx import Document
from docx.oxml.ns import qn

import upstream_verifier


TASK_ID = "cowork-evo/officeval-040-advanced-formatting-and-watermark"
OUTPUT_NAME = "简介素材_格式设置_水印版.docx"
SOURCE_DOC = "简介素材.docx"
SOURCE_IMAGE = "Image.png"
RUBRIC_VERSION = "omegause-officeval-040-adaptation-v0.1"
JUDGE_VERSION = "officeval-040-adaptation-v0.1"
SPECS = {
    "R001": ("指定 DOCX 可打开，标题与正文仍为可编辑文字", 3, "hurdle"),
    "R002": ("标题基础格式、发光和向上阴影全部正确", 5, "positive"),
    "R003": ("正文段落与字体格式正确，原始事实完整保留", 5, "hurdle"),
    "R004": ("第三个可见段的悬挂式两行首字下沉正确", 5, "positive"),
    "R005": ("全部名称替换正确，特殊文字六项格式完整", 4, "positive"),
    "R006": ("水印为真实校园图像、100% 缩放并位于背景层", 6, "hurdle"),
    "R007": ("水印具有 Office 冲蚀属性且不挤压正文", 4, "positive"),
    "R008": ("LibreOffice 可渲染且标题、专业列表和末段可提取", 3, "positive"),
    "R009": ("若正文被压平、原始事实丢失或水印变成正文图片，扣 10 分", -10, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {"criterion_id": identifier, "description": description, "weight": weight, "type": kind, "passed": bool(passed), "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def normalized_source_text(path: Path) -> str:
    text = "".join(paragraph.text.strip() for paragraph in Document(path).paragraphs)
    return text.replace("智能工程系", "智能软件工程系").replace("\t", "")


def normalized_output_text(path: Path) -> str:
    return "".join(paragraph.text.strip() for paragraph in Document(path).paragraphs).replace("\t", "")


def body_special_names_ok(path: Path) -> tuple[bool, str]:
    """只检查正文中的三处替换文字；标题不是 rubric 所说的“正文”。"""
    document = Document(path)
    matching_runs = []
    for paragraph in document.paragraphs[1:]:
        matching_runs.extend(run for run in paragraph.runs if run.text == "智能软件工程系")
    failures = []
    for index, run in enumerate(matching_runs, 1):
        rpr = run._r.get_or_add_rPr()
        east_asia = rpr.rFonts.get(qn("w:eastAsia")) if rpr.rFonts is not None else None
        emphasis = rpr.find(qn("w:em"))
        color = run.font.color.rgb
        if east_asia != "宋体" or run.font.size is None or abs(run.font.size.pt - 14) > 1e-6 or not run.bold or not run.italic or str(color).upper() != "00B050" or emphasis is None or emphasis.get(qn("w:val")) in {None, "none"}:
            failures.append(index)
    return len(matching_runs) == 3 and not failures, f"正文特殊名称={len(matching_runs)}处，格式失败={failures}"


def watermark_pixels_match(output: Path, source_image: Path) -> tuple[bool, str]:
    with zipfile.ZipFile(output) as archive:
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        payloads = [archive.read(name) for name in media]
        body_xml = archive.read("word/document.xml")
    with Image.open(source_image) as expected_image:
        expected_size = expected_image.size
        source_rgb = expected_image.convert("RGB")
        expected_pixels = Image.blend(source_rgb, Image.new("RGB", source_rgb.size, "white"), 0.78).convert("RGBA").tobytes()
    matching = 0
    for payload in payloads:
        try:
            with Image.open(io.BytesIO(payload)) as image:
                if image.size == expected_size and image.convert("RGBA").tobytes() == expected_pixels:
                    matching += 1
        except Exception:
            continue
    body_flattened = b"<w:drawing" in body_xml or b"<w:pict" in body_xml
    ok = matching == 1 and not body_flattened
    return ok, f"媒体数={len(media)}，像素级匹配确定性冲蚀结果={matching}，正文含图片={body_flattened}"


def render_check(path: Path) -> tuple[bool, str]:
    if not shutil.which("libreoffice") or not shutil.which("pdfinfo") or not shutil.which("pdftotext"):
        return False, "缺少渲染依赖"
    with tempfile.TemporaryDirectory(prefix="officeval040-render-") as tmp:
        proc = subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(path)], capture_output=True, text=True, timeout=180)
        pdf = Path(tmp) / f"{path.stem}.pdf"
        if proc.returncode != 0 or not pdf.is_file():
            return False, f"渲染失败：{proc.stderr[-300:]}"
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=30).stdout
        pages = next((int(line.split(":", 1)[1]) for line in info.splitlines() if line.startswith("Pages:")), 0)
        text = subprocess.run(["pdftotext", str(pdf), "-"], capture_output=True, text=True, timeout=60).stdout.replace("\n", "")
        ok = 1 <= pages <= 3 and all(value in text for value in ["智能软件工程系简介", "计算机应用技术", "面向岗位"])
        return ok, f"页数={pages}，关键文字可提取={ok}"


def evaluate(artifact_dir: Path, source_dir: Path) -> dict:
    output = artifact_dir / OUTPUT_NAME
    results = []
    try:
        document = Document(output)
        results.append(criterion("R001", bool(document.paragraphs), f"段落数={len(document.paragraphs)}，正文图片={len(document.inline_shapes)}"))
    except Exception as exc:
        results.append(criterion("R001", False, f"缺失或无法打开：{exc}"))
        return build_result(results, output)

    upstream = upstream_verifier.evaluate(str(artifact_dir))
    items = upstream.get("dim2_items", [])
    hits = [bool(item.get("hit")) for item in items]
    results.append(criterion("R002", len(hits) >= 2 and all(hits[:2]), f"标题两项={hits[:2]}"))

    content_ok = normalized_output_text(output) == normalized_source_text(source_dir / SOURCE_DOC)
    body_style_ok = len(hits) >= 4 and hits[2] and hits[3]
    results.append(criterion("R003", content_ok and body_style_ok, f"原文拼接一致={content_ok}，正文段落/字体={hits[2:4]}"))
    results.append(criterion("R004", len(hits) > 4 and hits[4], f"首字下沉上游命中={hits[4] if len(hits) > 4 else None}"))
    body_names_ok, body_names_evidence = body_special_names_ok(output)
    replace_ok = len(hits) > 5 and hits[5] and body_names_ok
    results.append(criterion("R005", replace_ok, f"无旧名称={hits[5] if len(hits) > 5 else None}；{body_names_evidence}；上游把标题误纳入正文特殊格式={hits[6] if len(hits) > 6 else None}"))

    pixel_ok, pixel_evidence = watermark_pixels_match(output, source_dir / SOURCE_IMAGE)
    scale_ok = len(hits) > 7 and hits[7]
    results.append(criterion("R006", pixel_ok and scale_ok, f"{pixel_evidence}；缩放背景={scale_ok}"))
    wash_range_ok = len(hits) > 9 and hits[8] and hits[9]
    results.append(criterion("R007", wash_range_ok, f"冲蚀/范围={hits[8:10] if len(hits) >= 10 else hits}"))
    render_pass, render_evidence = render_check(output)
    results.append(criterion("R008", render_pass, render_evidence))
    severe = not content_ok or not pixel_ok or bool(document.inline_shapes)
    results.append(criterion("R009", severe, f"事实/水印/压平严重错误={severe}，上游={upstream.get('total_score')}/{upstream.get('max_score')}"))
    return build_result(results, output)


def build_result(results: list[dict], output: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"]) + sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
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
    parser.add_argument("--source-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir, args.source_dir)
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
