#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook
from PIL import Image, ImageChops

TASK_ID = "cowork-evo/officeval-088-product-table"
OUTPUT = "officeval_088_product_table.xlsx"
EXPECTED = [
    {"key": "云溪", "source": "云溪5.0 D3滚筒洗衣机产品说明文档.docx", "name": "云溪5.0 D3滚筒洗衣机", "model": "云溪5.0 D3", "price": 4500, "crop": (70, 199, 202, 375), "ratio": 0.75},
    {"key": "统帅", "source": "统帅三桶懒人洗衣机产品说明文档.docx", "name": "统帅三桶懒人洗衣机", "model": "Leader L10510B12（标尺版）", "price": 3999, "crop": (232, 158, 471, 500), "ratio": 0.70},
    {"key": "小天鹅", "source": "小天鹅12KG滚筒洗衣机产品说明文档.docx", "name": "小天鹅12KG滚筒洗衣机", "model": "TG12VE40PRO", "price": 3039, "crop": (105, 367, 359, 669), "ratio": 0.84},
    {"key": "米家", "source": "米家8KG滚筒洗衣机产品说明文档.docx", "name": "米家8KG滚筒洗衣机", "model": "米家MG80V11D（滚筒款）", "price": 999, "crop": (141, 276, 315, 536), "ratio": 0.67},
    {"key": "美的", "source": "美的8KG波轮洗衣机产品说明文档.docx", "name": "美的8KG波轮洗衣机", "model": "美的MB80V36DT（波轮款）", "price": 846, "crop": (210, 120, 456, 498), "ratio": 0.65},
]
HASHES = {
    "云溪5.0 D3滚筒洗衣机产品说明文档.docx": "0f8ec77c24de31a8f685ab1089048605362e3b56e3cf823866d39006a860a8dc",
    "小天鹅12KG滚筒洗衣机产品说明文档.docx": "ad6e809a51bdc87fc667e1b52e2bc4e4560d547da57d835683a952beb32272dd",
    "模板.docx": "f3229e51c2c17c9a4dd8f4a13cb062625ca015448bfe865a51f8069377ecf782",
    "米家8KG滚筒洗衣机产品说明文档.docx": "e9cf7aff2d8de144d3b12c55af3416de4cb0a2d4246bfb499fd35894a583ed35",
    "统帅三桶懒人洗衣机产品说明文档.docx": "7c0cb03ac0c675531a3985a10694789f889e4907b7cef3032878915cec7504fa",
    "美的8KG波轮洗衣机产品说明文档.docx": "31c54e300fc61c34c0c1ee5fcc771b76e0f7eb11b46dd06c48f34bfeeacf2ef8",
    "选货汇总表_模板2_已填充 (1).xlsx": "ce9b586633945c4fee7b7d4d5c2ad149c2ed67a63591a4979f799bcfeb19724f",
}
WEIGHTS = {"R001": 10, "R002": 30, "R003": 20, "R004": 15, "R005": 10, "R006": 5, "R007": 4, "R008": 4, "R009": -20, "R010": -10}
KINDS = {"R001": "hurdle", "R002": "hurdle", "R003": "hurdle", "R004": "hurdle", "R005": "hurdle", "R006": "positive", "R007": "positive", "R008": "positive", "R009": "penalty", "R010": "penalty"}
CAPS = {"R001": 0.0, "R002": 0.45, "R003": 0.6, "R004": 0.7, "R005": 0.75}


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def criterion(key: str, score: float, evidence: str) -> dict[str, Any]:
    return {"criterion_id": key, "score": float(score), "evidence": evidence, "method": "deterministic", "weight": WEIGHTS[key], "type": KINDS[key]}


def first_media(path: Path) -> bytes:
    with zipfile.ZipFile(path) as archive:
        names = [name for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/")]
        if len(names) != 1:
            raise ValueError(f"{path.name}: expected one media item")
        return archive.read(names[0])


def expected_crops(root: Path) -> dict[str, Image.Image]:
    result = {}
    for item in EXPECTED:
        raw = first_media(root / "data/input_files" / item["source"])
        with tempfile.NamedTemporaryFile(suffix=".png") as file:
            file.write(raw)
            file.flush()
            with Image.open(file.name) as image:
                result[item["key"]] = image.convert("RGB").crop(item["crop"]).copy()
    return result


def recalc(path: Path) -> Path | None:
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice:
        return None
    directory = Path(tempfile.mkdtemp(prefix="officeval-088-recalc-"))
    profile = directory / "profile"
    completed = subprocess.run([libreoffice, "--headless", f"-env:UserInstallation={profile.resolve().as_uri()}", "--convert-to", "xlsx", "--outdir", str(directory), str(path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=240, check=False)
    result = directory / path.name
    return result if completed.returncode == 0 and result.is_file() else None


def render(path: Path) -> tuple[bool, str]:
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    pdfinfo = shutil.which("pdfinfo")
    if not libreoffice or not pdfinfo:
        return False, "缺少 LibreOffice 或 pdfinfo"
    directory = Path(tempfile.mkdtemp(prefix="officeval-088-render-"))
    profile = directory / "profile"
    completed = subprocess.run([libreoffice, "--headless", f"-env:UserInstallation={profile.resolve().as_uri()}", "--convert-to", "pdf", "--outdir", str(directory), str(path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=240, check=False)
    pdf = directory / f"{path.stem}.pdf"
    if completed.returncode != 0 or not pdf.is_file() or pdf.stat().st_size < 1000:
        return False, completed.stdout[-400:]
    info = subprocess.run([pdfinfo, str(pdf)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=False).stdout
    match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
    return bool(match and int(match.group(1)) > 0), f"PDF={pdf.stat().st_size} bytes; {match.group(0) if match else 'no pages'}"


def build_result(results: list[dict[str, Any]], output: Path) -> dict[str, Any]:
    positive_total = sum(value for value in WEIGHTS.values() if value > 0)
    earned = sum(WEIGHTS[item["criterion_id"]] * item["score"] for item in results if WEIGHTS[item["criterion_id"]] > 0)
    penalties = sum(WEIGHTS[item["criterion_id"]] for item in results if WEIGHTS[item["criterion_id"]] < 0 and item["score"] >= 1)
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    failures = [item["criterion_id"] for item in results if KINDS[item["criterion_id"]] == "hurdle" and item["score"] < 1]
    if failures:
        reward = min(reward, min(CAPS[key] for key in failures))
    return {"task_id": TASK_ID, "rubric_version": "officeval-088-xlsx-v0.2", "judge_version": "2.0.0", "artifact_sha256": [digest(output)] if output.is_file() else [], "criterion_results": [{"criterion_id": item["criterion_id"], "score": item["score"], "evidence": item["evidence"], "method": item["method"]} for item in sorted(results, key=lambda x: x["criterion_id"])], "reward": round(reward, 6), "passed": reward >= 0.8 and not failures, "hurdle_failures": failures}


def evaluate(artifact_dir: Path, do_render: bool = True) -> dict[str, Any]:
    root = Path(__file__).parents[1]
    input_dir = root / "data/input_files"
    output = artifact_dir / OUTPUT
    names = sorted(path.name for path in artifact_dir.iterdir()) if artifact_dir.is_dir() else []
    results: list[dict[str, Any]] = []
    valid = output.is_file() and output.stat().st_size > 0 and zipfile.is_zipfile(output)
    if valid:
        try:
            wb = load_workbook(output, data_only=False)
            valid = wb.sheetnames == ["Products", "Top3 Analysis"]
            wb.close()
        except Exception:
            valid = False
    results.append(criterion("R001", 1 if valid else 0, "XLSX 可解析且包含 Products/Top3 Analysis" if valid else "缺少或损坏输出"))
    if not valid:
        results.extend([criterion("R002", 0, "文件无效"), criterion("R003", 0, "文件无效"), criterion("R004", 0, "文件无效"), criterion("R005", 0, "文件无效"), criterion("R006", 0, "文件无效"), criterion("R007", 1 if all(digest(input_dir / name) == value for name, value in HASHES.items()) else 0, "输入哈希"), criterion("R008", 0, f"目录={names}"), criterion("R009", 1, "无效输出"), criterion("R010", 1 if any(Path(n).suffix.lower() != ".xlsx" for n in names) else 0, f"目录={names}")])
        return build_result(results, output)

    workbook = load_workbook(output, data_only=False)
    products = workbook["Products"]
    summary = workbook["Top3 Analysis"]
    headers_ok = [products.cell(1, col).value for col in range(1, 7)] == ["产品名称", "核心参数", "适用场景", "指导价格", "展示图片", "来源文档"]
    row_issues = []
    for row, item in enumerate(EXPECTED, 2):
        values = " ".join(str(products.cell(row, col).value or "") for col in range(1, 7))
        for required in (item["name"], item["model"], item["source"]):
            if required not in values and required.replace(" ", "") not in values.replace(" ", ""):
                row_issues.append(f"{item['key']}:missing {required}")
        if products.cell(row, 4).value != item["price"]:
            row_issues.append(f"{item['key']}:price={products.cell(row, 4).value}")
    ranked = sorted(EXPECTED, key=lambda item: item["price"], reverse=True)[:3]
    summary_formulas_ok = True
    for row, item in enumerate(ranked, 2):
        source_row = EXPECTED.index(item) + 2
        summary_formulas_ok &= summary.cell(row, 2).value == f"=Products!A{source_row}"
        summary_formulas_ok &= summary.cell(row, 4).value == f"=Products!D{source_row}"
        summary_formulas_ok &= summary.cell(row, 5).value == f"=Products!F{source_row}"
    recalc_path = recalc(output)
    cross_ok = False
    if recalc_path:
        cached = load_workbook(recalc_path, data_only=True, read_only=True)
        cs = cached["Top3 Analysis"]
        cross_ok = all(cs.cell(row, 2).value == item["name"] and float(cs.cell(row, 4).value) == item["price"] for row, item in enumerate(ranked, 2))
        cached.close()
    fact_score = 1.0 if headers_ok and not row_issues and summary_formulas_ok and cross_ok else 0.0
    results.append(criterion("R002", fact_score, f"headers={headers_ok}; issues={row_issues[:5]}; top3_formulas={summary_formulas_ok}; cross_sheet={cross_ok}"))

    crops = expected_crops(root)
    image_issues = []
    if len(products._images) != 5:
        image_issues.append(f"image_count={len(products._images)}")
    for image in products._images:
        anchor = getattr(image.anchor, "_from", None)
        if anchor is None or anchor.col != 4 or not 1 <= anchor.row <= 5:
            image_issues.append("invalid image anchor")
            continue
        key = EXPECTED[anchor.row - 1]["key"]
        with tempfile.NamedTemporaryFile(suffix=".png") as file:
            file.write(image._data()); file.flush()
            with Image.open(file.name) as actual:
                actual = actual.convert("RGB")
                expected_image = crops[key]
                if actual.size != expected_image.size:
                    image_issues.append(f"{key}:size={actual.size}")
                else:
                    diff = ImageChops.difference(actual, expected_image)
                    extrema = max(channel[1] for channel in diff.getextrema())
                    if extrema > 8:
                        image_issues.append(f"{key}:pixel_diff={extrema}")
    results.append(criterion("R003", 1 if not image_issues else 0, f"images={image_issues or '5 anchors and crops match'}"))
    calc = workbook.calculation
    native_formulas = all(isinstance(summary.cell(row, col).value, str) and summary.cell(row, col).value.startswith("=") for row in range(2, 5) for col in (2, 3, 4, 5))
    calc_ok = getattr(calc, "calcMode", None) == "auto" and bool(getattr(calc, "fullCalcOnLoad", False)) and bool(getattr(calc, "forceFullCalc", False))
    results.append(criterion("R004", 1 if native_formulas and calc_ok else 0, f"native_formulas={native_formulas}; calc={calc_ok}"))
    structure_ok = products.max_row == 6 and products.max_column == 6 and summary.max_row == 4 and summary.max_column == 5 and products.freeze_panes == "A2" and summary.freeze_panes == "A2"
    results.append(criterion("R005", 1 if structure_ok else 0, f"Products={products.max_row}x{products.max_column}; Summary={summary.max_row}x{summary.max_column}; freeze={products.freeze_panes}/{summary.freeze_panes}"))
    hash_ok = all(digest(input_dir / name) == value for name, value in HASHES.items())
    results.append(criterion("R006", 1 if hash_ok else 0, f"input_hashes={hash_ok}"))
    results.append(criterion("R007", 1 if hash_ok else 0, "所有七个输入文件与已审计清单哈希一致"))
    render_ok, render_evidence = render(output) if do_render else (True, "跳过渲染")
    results.append(criterion("R008", 1 if render_ok else 0, render_evidence))
    results.append(criterion("R009", 1 if digest(output) == digest(input_dir / EXCEL_TEMPLATE) else 0, f"unchanged={digest(output) == digest(input_dir / EXCEL_TEMPLATE)}"))
    results.append(criterion("R010", 1 if any(Path(n).suffix.lower() != ".xlsx" for n in names) else 0, f"目录={names}"))
    workbook.close()
    return build_result(results, output)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    parser.add_argument("--skip-render", action="store_true")
    args = parser.parse_args()
    result = evaluate(args.artifact_dir.resolve(), not args.skip_render)
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True); args.result.write_text(payload, encoding="utf-8")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True); args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    print(payload, end="")
    raise SystemExit(0 if result["passed"] else 1)


EXCEL_TEMPLATE = "选货汇总表_模板2_已填充 (1).xlsx"
if __name__ == "__main__":
    main()
