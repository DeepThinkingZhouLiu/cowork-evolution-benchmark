#!/usr/bin/env python3
"""officeval_036 的 Harbor Judge。

它将 PDF 人工转录的 48 行真实记录与 Word 模板原有 2 行分开记录，
并复用上游 OOXML 格式检查，再补充严格顺序、行高、模板保留和实际渲染。
"""

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

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE

import upstream_verifier


OUTPUT_NAME = "public_ready_statement_收支记录整理.docx"
INPUT_NAME = "public_ready_statement.docx"
TASK_ID = "cowork-evo/officeval-036-pdf-statement-to-word-table"
RUBRIC_VERSION = "omegause-officeval-036-adaptation-v0.1"
JUDGE_VERSION = "officeval-036-adaptation-v0.1"
EXPECTED_ROW_HEIGHT_CM = 1.80

CRITERION_SPECS = {
    "R001": ("指定 DOCX 可打开，且收支记录为 5 列可编辑 Word 表格", 2, "hurdle"),
    "R002": ("2 行模板原记录与 PDF 48 行记录完整、准确、按顺序录入", 8, "positive"),
    "R003": ("50 个数据行使用完全一致的固定行高且不跨页拆分", 3, "positive"),
    "R004": ("表头、对齐和 3.68 厘米统一列宽满足规范", 3, "positive"),
    "R005": ("数据字体与单元格段落格式满足规范", 2, "positive"),
    "R006": ("横向边框、美元金额格式和收支空白列正确", 3, "positive"),
    "R007": ("跨页时自动重复表头且样式一致", 2, "positive"),
    "R008": ("模板图片、标题和第二个空续表表头完整保留", 2, "positive"),
    "R009": ("LibreOffice 实际渲染后 50 行关键文字均可读且无行裁切", 3, "positive"),
    "R010": ("5 行及以上记录缺失、重复、错序或金额错列", -8, "penalty"),
    "R011": ("数据行高不一致、跨页拆分或实际渲染缺文", -5, "penalty"),
    "R012": ("收支记录被截图或非编辑对象替代", -8, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str = "") -> dict:
    description, weight, kind = CRITERION_SPECS[identifier]
    return {
        "criterion_id": identifier,
        "description": description,
        "weight": weight,
        "type": kind,
        "passed": bool(passed),
        "score": 1.0 if passed else 0.0,
        "evidence": evidence,
        "method": "deterministic",
    }


def normalized(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def media_hashes(path: Path) -> dict[str, str]:
    hashes = {}
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith("word/media/") and not name.endswith("/"):
                hashes[name] = hashlib.sha256(archive.read(name)).hexdigest()
    return hashes


def record_mismatches(actual: list[tuple[str, ...]], expected: list[list[str]]) -> list[str]:
    errors = []
    if len(actual) != len(expected):
        errors.append(f"数据行数={len(actual)}，期望={len(expected)}")
    if len(actual) < len(expected):
        for index in range(len(actual), len(expected)):
            errors.append(f"第{index + 1}行缺失：期望={expected[index]!r}")
    elif len(actual) > len(expected):
        for index in range(len(expected), len(actual)):
            errors.append(f"第{index + 1}行为多余记录：实际={list(actual[index])!r}")
    for index, (actual_row, expected_row) in enumerate(zip(actual, expected), start=1):
        if list(actual_row) != expected_row:
            errors.append(f"第{index}行：期望={expected_row!r}，实际={list(actual_row)!r}")
    return errors


def row_height_check(document: Document) -> tuple[bool, list[str]]:
    table = document.tables[1]
    issues = []
    heights = []
    for index, row in enumerate(table.rows[1:], start=1):
        if row.height is None:
            issues.append(f"第{index}行未设置固定行高")
            continue
        height_cm = row.height.cm
        heights.append(round(height_cm, 4))
        if abs(height_cm - EXPECTED_ROW_HEIGHT_CM) > 0.03:
            issues.append(f"第{index}行高={height_cm:.3f}cm")
        if row.height_rule != WD_ROW_HEIGHT_RULE.EXACTLY:
            issues.append(f"第{index}行高规则={row.height_rule}")
        tr_pr = row._tr.trPr
        if tr_pr is None or tr_pr.find(upstream_verifier.qn("w:cantSplit")) is None:
            issues.append(f"第{index}行未禁止跨页拆分")
    if len(set(heights)) > 1:
        issues.append(f"行高不一致：{sorted(set(heights))}")
    return not issues and len(heights) == 50, issues


def template_preservation_check(source_path: Path, output_path: Path, document: Document, points) -> tuple[bool, list[str]]:
    issues = []
    if media_hashes(source_path) != media_hashes(output_path):
        issues.append("模板图片集合或图片内容改变")
    all_text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    if "Activity details" not in all_text:
        issues.append("缺少 Activity details 标题")
    if len(document.tables) != 3:
        issues.append(f"模板表格数={len(document.tables)}，期望=3")
    elif [cell.text.strip() for cell in document.tables[2].rows[0].cells] != upstream_verifier.HEADER:
        issues.append("第二个空续表表头改变")
    if points[8].hit:
        issues.append(points[8].detail or "封面图片尺寸不符")
    if points[9].hit:
        issues.append(points[9].detail or "标题变成艺术字")
    return not issues, issues


def render_check(path: Path, expected: list[list[str]]) -> tuple[bool, str]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    if libreoffice is None or pdftotext is None:
        return False, "缺少 libreoffice 或 pdftotext"
    with tempfile.TemporaryDirectory(prefix="cowork-evo-036-") as temp_dir:
        temp_root = Path(temp_dir)
        profile_uri = (temp_root / "lo-profile").resolve().as_uri()
        converted = subprocess.run(
            [
                libreoffice,
                "--headless",
                f"-env:UserInstallation={profile_uri}",
                "--convert-to",
                "pdf",
                "--outdir",
                temp_dir,
                str(path),
            ],
            capture_output=True,
            text=True,
            timeout=300,
            check=False,
        )
        pdf_path = temp_root / f"{path.stem}.pdf"
        if converted.returncode != 0 or not pdf_path.is_file():
            return False, (converted.stderr or converted.stdout or "PDF 转换失败")[-1000:]
        # -raw 会保留 Word 表格单元格的文本流顺序；-layout 会把
        # 同一行其他金额列插入多行 Description 中间，不适合做裁切检查。
        extracted = subprocess.run(
            [pdftotext, "-raw", str(pdf_path), "-"],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if extracted.returncode != 0:
            return False, (extracted.stderr or "PDF 文字抽取失败")[-1000:]
        pages = [page for page in extracted.stdout.split("\f") if page.strip()]
        rendered = normalized("".join(pages))
    required = sorted({normalized(cell) for row in expected for cell in row if cell})
    missing = [item for item in required if item not in rendered]
    page_ok = 5 <= len(pages) <= 8
    return page_ok and not missing, f"页数={len(pages)}，缺失字段={missing[:20]}"


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    expected_path = Path(__file__).with_name("expected-records.json")
    expected = json.loads(expected_path.read_text(encoding="utf-8"))["records"]
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)

    ctx, load_error = upstream_verifier.load_docx(output_path)
    if ctx is None:
        results.append(criterion("R001", False, load_error))
        return build_result(results, output_path)
    try:
        table_xml = upstream_verifier.find_record_table(ctx)
        if table_xml is None:
            results.append(criterion("R001", False, "未找到指定 5 列收支记录表"))
            return build_result(results, output_path)
        points = upstream_verifier.dimension2(ctx, table_xml)
        editable_ok = points[0].hit
        results.append(criterion("R001", editable_ok, points[0].detail))

        actual_records = upstream_verifier.data_records(table_xml)
        mismatches = record_mismatches(actual_records, expected)
        results.append(
            criterion(
                "R002",
                not mismatches,
                "50 行记录逐字段与 PDF/模板一致" if not mismatches else "；".join(mismatches[:12]),
            )
        )

        document = Document(output_path)
        heights_ok, height_issues = row_height_check(document)
        results.append(criterion("R003", heights_ok, "50 行均为 1.80cm 且禁止跨页拆分" if heights_ok else "；".join(height_issues[:20])))
        results.append(criterion("R004", points[0].hit and points[1].hit, f"{points[0].detail}；{points[1].detail}"))
        results.append(criterion("R005", points[2].hit and points[3].hit, f"{points[2].detail}；{points[3].detail}"))
        amount_ok = points[4].hit and points[5].hit and not points[10].hit
        results.append(criterion("R006", amount_ok, f"{points[4].detail}；{points[5].detail}；{points[10].detail}"))
        results.append(criterion("R007", points[7].hit, points[7].detail))

        source_path = Path(__file__).resolve().parents[1] / "data" / "input_files" / INPUT_NAME
        template_ok, template_issues = template_preservation_check(source_path, output_path, document, points)
        results.append(criterion("R008", template_ok, "模板图片、标题与续表表头完整" if template_ok else "；".join(template_issues)))

        rendered, render_evidence = render_check(output_path, expected)
        results.append(criterion("R009", rendered, render_evidence))
        results.append(criterion("R010", len(mismatches) >= 5, f"记录错误数={len(mismatches)}"))
        results.append(criterion("R011", not heights_ok or not rendered, f"行高通过={heights_ok}，渲染通过={rendered}"))
        results.append(criterion("R012", not editable_ok, points[0].detail))
        return build_result(results, output_path)
    finally:
        ctx.zip_file.close()


def build_result(results: list[dict], artifact_path: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "因交付不存在、损坏或门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])

    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = 0.0
    digest = hashlib.sha256(artifact_path.read_bytes()).hexdigest() if artifact_path.is_file() else hashlib.sha256(b"").hexdigest()
    normalized_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        normalized_results.append(
            {
                "criterion_id": item["criterion_id"],
                "score": score,
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
        "passed": reward >= 0.75,
        "hurdle_failures": hurdle_failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
