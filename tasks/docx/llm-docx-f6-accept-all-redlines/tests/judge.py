#!/usr/bin/env python3
"""LLM-DOCX-Editing F6 的确定性 Harbor Judge。"""

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
from lxml import etree


OUTPUT_NAME = "contract-accepted.docx"
TASK_ID = "cowork-evo/llm-docx-f6-accept-all-redlines"
RUBRIC_VERSION = "llm-docx-editing-f6-adaptation-v0.1"
JUDGE_VERSION = "llm-docx-editing-f6-adaptation-v0.1"
HURDLE_CAPS = {"R001": 0.0, "R006": 0.74, "R007": 0.74}

CANONICAL_PARAGRAPHS = [
    "Supply Agreement",
    'This Supply Agreement is entered into between Global Manufacturing Corp ("Supplier") and TechBuild Inc. ("Buyer").',
    "1. Delivery Terms",
    "Supplier shall deliver all goods within forty-five (45) calendar days of receipt of a purchase order.",
    "2. Warranty",
    "Supplier warrants that all goods shall be free from defects for a period of twenty-four (24) months from the date of delivery.",
    "3. Price Adjustment",
    "Supplier may request a price adjustment no more than once per calendar year, with at least ninety (90) days written notice. Any increase shall not exceed the greater of 3% or the annual CPI change.",
    "4. Termination",
    "Either party may terminate with ninety (90) days written notice.",
]
HEADINGS = [
    "1. Delivery Terms",
    "2. Warranty",
    "3. Price Adjustment",
    "4. Termination",
]
ACCEPTED_TERMS = [
    "Global Manufacturing Corp",
    "forty-five (45)",
    "twenty-four (24) months",
    "ninety (90) days",
]
STALE_TERMS = [
    "Acme Manufacturing",
    "thirty (30)",
    "twelve (12) months",
    "sixty (60)",
]
REVISION_NAMES = {
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
    "pPrChange",
    "rPrChange",
    "tblPrChange",
    "tblGridChange",
    "trPrChange",
    "tcPrChange",
    "sectPrChange",
    "numberingChange",
    "trackRevisions",
}

CRITERION_SPECS = {
    "R001": ("指定 DOCX 可打开、可编辑并包含完整 Supply Agreement", 3, "hurdle"),
    "R002": ("供应商名称和交付期均为接受后的新值", 3, "positive"),
    "R003": ("质保期和终止通知期均为接受后的新值", 3, "positive"),
    "R004": ("整段 Price Adjustment 插入条款完整保留", 4, "positive"),
    "R005": ("所有旧值在可见正文和底层 XML 中均已删除", 3, "positive"),
    "R006": ("所有插入、删除、移动和属性修订 OOXML 节点均已清除", 4, "hurdle"),
    "R007": ("清洁文档的 10 个段落逐项等于原文件的接受视图", 5, "hurdle"),
    "R008": ("四个编号章节的顺序和可编辑结构保持", 3, "positive"),
    "R009": ("DOCX 可真实渲染为清洁、无红线且可读的一页合同", 3, "positive"),
    "R010": ("任一旧值仍作为可见或底层修订内容残留", -8, "penalty"),
    "R011": ("任一修订节点、红线标记或未接受变更残留", -8, "penalty"),
    "R012": ("正文截图化、明显截断或无法渲染", -10, "penalty"),
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
    return re.sub(r"\s+", " ", text or "").strip().lower()


def inspect_ooxml(path: Path) -> tuple[dict[str, int], str, int]:
    counts = {name: 0 for name in sorted(REVISION_NAMES)}
    xml_text_parts = []
    with zipfile.ZipFile(path) as archive:
        media_count = sum(
            1
            for name in archive.namelist()
            if name.startswith("word/media/") and not name.endswith("/")
        )
        for name in archive.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = etree.fromstring(archive.read(name))
            for element in root.iter():
                local = etree.QName(element).localname
                if local in counts:
                    counts[local] += 1
            xml_text_parts.extend(root.itertext())
    return counts, " ".join(xml_text_parts), media_count


def render_check(path: Path) -> tuple[bool, str, int]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    if libreoffice is None or pdftotext is None:
        return False, "缺少 libreoffice 或 pdftotext", 0
    with tempfile.TemporaryDirectory(prefix="cowork-evo-f6-") as temp_dir:
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
            return False, (converted.stderr or converted.stdout or "PDF 转换失败")[-1000:], 0
        extracted = subprocess.run(
            [pdftotext, "-raw", str(pdf_path), "-"],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if extracted.returncode != 0:
            return False, (extracted.stderr or "PDF 文字抽取失败")[-1000:], 0
        pages = [page for page in extracted.stdout.split("\f") if page.strip()]
        rendered = normalized("\n".join(pages))

    required = [
        "Supply Agreement",
        *HEADINGS,
        *ACCEPTED_TERMS,
        "once per calendar year",
        "greater of 3% or the annual CPI change",
    ]
    missing = [item for item in required if normalized(item) not in rendered]
    stale = [item for item in STALE_TERMS if normalized(item) in rendered]
    ok = len(pages) == 1 and not missing and not stale and len(rendered) >= 650
    return ok, f"页数={len(pages)}，渲染缺失={missing}，渲染旧值={stale}，字符={len(rendered)}", len(pages)


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)

    try:
        document = Document(output_path)
        paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
        full_text = "\n".join(paragraphs)
        revision_counts, xml_text, media_count = inspect_ooxml(output_path)
    except (OSError, ValueError, zipfile.BadZipFile, etree.XMLSyntaxError) as error:
        results.append(criterion("R001", False, f"DOCX 无法读取：{error}"))
        return build_result(results, output_path)

    editable_ok = (
        len(paragraphs) >= 10
        and len(normalized(full_text)) >= 650
        and media_count == 0
        and paragraphs[0] == "Supply Agreement"
    )
    results.append(
        criterion(
            "R001",
            editable_ok,
            f"段落={len(paragraphs)}，文本字符={len(normalized(full_text))}，媒体={media_count}，首段={paragraphs[0] if paragraphs else ''}",
        )
    )

    visible = normalized(full_text)
    party_delivery = all(normalized(term) in visible for term in ACCEPTED_TERMS[:2])
    results.append(criterion("R002", party_delivery, f"新供应商={'global manufacturing corp' in visible}，新交付期={'forty-five (45)' in visible}"))

    warranty_termination = all(normalized(term) in visible for term in ACCEPTED_TERMS[2:])
    termination_exact = normalized(CANONICAL_PARAGRAPHS[9]) in visible
    results.append(criterion("R003", warranty_termination and termination_exact, f"新质保={'twenty-four (24) months' in visible}，新通知期={'ninety (90) days' in visible}，终止句={termination_exact}"))

    price_terms = [
        "no more than once per calendar year",
        "at least ninety (90) days written notice",
        "greater of 3% or the annual cpi change",
    ]
    price_ok = normalized(CANONICAL_PARAGRAPHS[7]) in visible and all(term in visible for term in price_terms)
    results.append(criterion("R004", price_ok, f"价格调整完整句={normalized(CANONICAL_PARAGRAPHS[7]) in visible}，子要素={[term in visible for term in price_terms]}"))

    combined = normalized(full_text + " " + xml_text)
    stale_found = [term for term in STALE_TERMS if normalized(term) in combined]
    results.append(criterion("R005", not stale_found, f"可见正文+XML 旧值={stale_found}"))

    revision_found = {name: count for name, count in revision_counts.items() if count}
    results.append(criterion("R006", not revision_found, f"残留修订节点={revision_found}"))

    exact_ok = paragraphs == CANONICAL_PARAGRAPHS
    mismatches = []
    for index in range(max(len(paragraphs), len(CANONICAL_PARAGRAPHS))):
        actual = paragraphs[index] if index < len(paragraphs) else "<缺失>"
        expected = CANONICAL_PARAGRAPHS[index] if index < len(CANONICAL_PARAGRAPHS) else "<多余>"
        if actual != expected:
            mismatches.append({"index": index, "expected": expected, "actual": actual})
    results.append(criterion("R007", exact_ok, f"段落={len(paragraphs)}/10，差异={mismatches[:3]}"))

    heading_positions = []
    for heading in HEADINGS:
        try:
            heading_positions.append(paragraphs.index(heading))
        except ValueError:
            heading_positions.append(-1)
    structure_ok = heading_positions == [2, 4, 6, 8] and media_count == 0
    results.append(criterion("R008", structure_ok, f"章节位置={dict(zip(HEADINGS, heading_positions))}，媒体={media_count}"))

    render_ok, render_evidence, _ = render_check(output_path)
    results.append(criterion("R009", render_ok, render_evidence))

    results.append(criterion("R010", bool(stale_found), f"旧值={stale_found}"))
    results.append(criterion("R011", bool(revision_found), f"修订节点={revision_found}"))
    unusable = media_count > 0 or len(normalized(full_text)) < 500 or not render_ok
    results.append(criterion("R012", unusable, f"媒体={media_count}，文本字符={len(normalized(full_text))}，渲染通过={render_ok}"))
    return build_result(results, output_path)


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
    hurdle_failures = [
        item["criterion_id"]
        for item in results
        if item["type"] == "hurdle" and not item["passed"]
    ]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[identifier] for identifier in hurdle_failures))
    digest = hashlib.sha256(artifact_path.read_bytes()).hexdigest() if artifact_path.is_file() else hashlib.sha256(b"").hexdigest()

    criterion_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        criterion_results.append(
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
        "criterion_results": criterion_results,
        "reward": round(reward, 6),
        "passed": reward >= 0.75,
        "hurdle_failures": hurdle_failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
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
