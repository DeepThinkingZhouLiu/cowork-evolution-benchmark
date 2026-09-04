#!/usr/bin/env python3
"""LLM-DOCX-Editing S4 的确定性 Harbor Judge。"""

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


OUTPUT_NAME = "legal-memo-revised.docx"
TASK_ID = "cowork-evo/llm-docx-s4-revise-memo-meeting-notes"
RUBRIC_VERSION = "llm-docx-editing-s4-adaptation-v0.1"
JUDGE_VERSION = "llm-docx-editing-s4-adaptation-v0.1"

CRITERION_SPECS = {
    "R001": ("指定 DOCX 可打开、含可编辑正文并保留备忘录结构", 3, "hurdle"),
    "R002": ("摘要与建议均把 Option B 作为首选并明确两个监管机构", 5, "positive"),
    "R003": ("所有目标时间更新为 Q3 2026 且无 Q4 2025", 3, "positive"),
    "R004": ("预算上限与策略预算正确且旧超额区间已移除", 4, "positive"),
    "R005": ("Option B 章节实质加强监管先例和执法趋势", 4, "positive"),
    "R006": ("Option A 章节和法律分析保留并定位为 fallback", 4, "positive"),
    "R007": ("未要求修改的关键背景与法律事实完整保留", 4, "positive"),
    "R008": ("头部、六个章节和必要分析保持完整", 3, "positive"),
    "R009": ("实际渲染可读且所有章节与新决策可见", 3, "positive"),
    "R010": ("残留旧首选、旧时间或旧超额预算", -8, "penalty"),
    "R011": ("三个以上关键事实或 Option A 主要分析丢失", -7, "penalty"),
    "R012": ("正文截图化、明显截断或无法渲染", -10, "penalty"),
}

HEADINGS = [
    "I. Executive Summary",
    "II. Background",
    "III. Option A: Federal Litigation",
    "IV. Option B: Regulatory Filing",
    "V. Comparative Analysis",
    "VI. Recommendation",
]

PRESERVED_ANCHORS = [
    "Westfield Technologies Inc.",
    "January 15, 2026",
    "2.3 million customer records",
    "compromised API endpoint",
    "August 2024",
    "$3.2 million",
    "$5 million",
    "Van Buren v. United States",
    "City of Santa Barbara v. Superior Court",
    "California's Unfair Competition Law",
]


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


def section_text(paragraphs: list[str], start: str, end: str | None) -> str:
    try:
        start_index = paragraphs.index(start)
    except ValueError:
        return ""
    if end is None:
        end_index = len(paragraphs)
    else:
        try:
            end_index = paragraphs.index(end, start_index + 1)
        except ValueError:
            end_index = len(paragraphs)
    return "\n".join(paragraphs[start_index + 1 : end_index])


def count_media(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(
            1
            for name in archive.namelist()
            if name.startswith("word/media/") and not name.endswith("/")
        )


def render_check(path: Path) -> tuple[bool, str, int]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    if libreoffice is None or pdftotext is None:
        return False, "缺少 libreoffice 或 pdftotext", 0
    with tempfile.TemporaryDirectory(prefix="cowork-evo-s4-") as temp_dir:
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
        *HEADINGS,
        "Option B",
        "Q3 2026",
        "$500,000",
        "fallback",
    ]
    missing = [item for item in required if normalized(item) not in rendered]
    ok = 2 <= len(pages) <= 4 and not missing and len(rendered) >= 3500
    return ok, f"页数={len(pages)}，渲染缺失={missing}，文本字符={len(rendered)}", len(pages)


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)

    try:
        document = Document(output_path)
        paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs]
        text = "\n".join(paragraphs)
        text_norm = normalized(text)
        media_count = count_media(output_path)
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        results.append(criterion("R001", False, f"DOCX 无法读取：{error}"))
        return build_result(results, output_path)

    nonempty = [paragraph for paragraph in paragraphs if paragraph]
    header_ok = all(
        marker in text
        for marker in [
            "MEMORANDUM",
            "TO: Sarah Chen, General Counsel",
            "FROM: David Nakamura, Senior Associate",
            "DATE: March 10, 2026",
        ]
    )
    structure_ok = len(paragraphs) >= 25 and len(nonempty) >= 25 and header_ok
    results.append(
        criterion(
            "R001",
            structure_ok and len(text_norm) >= 3000,
            f"段落={len(paragraphs)}，非空段落={len(nonempty)}，文本字符={len(text_norm)}，头部完整={header_ok}",
        )
    )

    executive = normalized(section_text(paragraphs, HEADINGS[0], HEADINGS[1]))
    recommendation = normalized(section_text(paragraphs, HEADINGS[5], None))
    option_b_primary = all(
        term in executive and term in recommendation
        for term in ["option b", "regulatory"]
    )
    regulators = all(
        term in executive + "\n" + recommendation
        for term in ["federal trade commission", "california attorney general"]
    ) or all(term in executive + "\n" + recommendation for term in ["ftc", "california attorney general"])
    recommendation_primary = "option b" in recommendation and (
        "primary enforcement strategy" in recommendation
        or "recommend pursuing option b" in recommendation
    )
    results.append(
        criterion(
            "R002",
            option_b_primary and regulators and recommendation_primary,
            f"摘要同步={option_b_primary}，监管机构={regulators}，最终首选={recommendation_primary}",
        )
    )

    q3_count = text_norm.count("q3 2026")
    q4_count = text_norm.count("q4 2025")
    timeline_ok = q3_count >= 2 and q4_count == 0
    results.append(criterion("R003", timeline_ok, f"Q3 2026 次数={q3_count}，Q4 2025 次数={q4_count}"))

    ceiling_present = "$500,000" in text
    stale_amounts = [amount for amount in ["$750,000", "$1,200,000", "$850,000", "$1,350,000"] if amount in text]
    current_range_present = "$200,000" in text and "$350,000" in text
    budget_ok = ceiling_present and current_range_present and not stale_amounts
    results.append(
        criterion(
            "R004",
            budget_ok,
            f"预算上限={ceiling_present}，上限内区间={current_range_present}，旧金额={stale_amounts}",
        )
    )

    option_b = normalized(section_text(paragraphs, HEADINGS[3], HEADINGS[4]))
    option_b_concepts = {
        "enforcement": "enforcement" in option_b,
        "trend_or_precedent": "trend" in option_b or "precedent" in option_b,
        "section_5": "section 5" in option_b,
        "ccpa": "ccpa" in option_b,
        "settlement_or_penalty": "settlement" in option_b or "penalt" in option_b,
    }
    strengthened = all(option_b_concepts.values()) and len(option_b) >= 900
    results.append(criterion("R005", strengthened, f"概念={option_b_concepts}，章节字符={len(option_b)}"))

    option_a = normalized(section_text(paragraphs, HEADINGS[2], HEADINGS[3]))
    option_a_anchors = ["cfaa", "van buren", "$5 million", "city of santa barbara", "ccpa"]
    # 原始 S4 的 Option A 正文规范化后为 1098 个字符；1000 的下限允许
    # 轻微措辞调整，但会拒绝把该章节压成一句空泛的 fallback 说明。
    preserved_option_a = all(anchor in option_a for anchor in option_a_anchors) and len(option_a) >= 1000
    fallback_count = text_norm.count("fallback")
    fallback_ok = fallback_count >= 2 and "federal litigation" in text_norm
    results.append(
        criterion(
            "R006",
            preserved_option_a and fallback_ok,
            f"Option A 锚点保留={preserved_option_a}，fallback 次数={fallback_count}",
        )
    )

    missing_anchors = [anchor for anchor in PRESERVED_ANCHORS if normalized(anchor) not in text_norm]
    results.append(
        criterion(
            "R007",
            not missing_anchors,
            "10 个关键事实全部保留" if not missing_anchors else f"缺失={missing_anchors}",
        )
    )

    missing_headings = [heading for heading in HEADINGS if heading not in paragraphs]
    structure_preserved = header_ok and not missing_headings and len(nonempty) >= 25 and len(text_norm) >= 4000
    results.append(
        criterion(
            "R008",
            structure_preserved,
            f"缺失章节={missing_headings}，非空段落={len(nonempty)}，文本字符={len(text_norm)}",
        )
    )

    render_ok, render_evidence, _ = render_check(output_path)
    results.append(criterion("R009", render_ok, render_evidence))

    stale_primary = bool(
        re.search(r"recommend(?:ed|ing)?\s+(?:that\s+we\s+)?pursu(?:e|ing)\s+option\s+a", text_norm)
        or "option a (federal litigation) as the primary" in text_norm
    )
    stale_decision = stale_primary or q4_count > 0 or bool(stale_amounts)
    results.append(
        criterion(
            "R010",
            stale_decision,
            f"旧首选={stale_primary}，旧时间={q4_count}，旧预算={stale_amounts}",
        )
    )

    severe_loss = len(missing_anchors) >= 3 or not preserved_option_a or len(nonempty) < 22
    results.append(
        criterion(
            "R011",
            severe_loss,
            f"缺失关键事实={len(missing_anchors)}，Option A 保留={preserved_option_a}，非空段落={len(nonempty)}",
        )
    )

    rasterized_or_broken = media_count > 0 and len(text_norm) < 3000
    unusable = rasterized_or_broken or len(text_norm) < 2000 or not render_ok
    results.append(
        criterion(
            "R012",
            unusable,
            f"媒体文件={media_count}，文本字符={len(text_norm)}，渲染通过={render_ok}",
        )
    )
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
        reward = 0.0
    digest = (
        hashlib.sha256(artifact_path.read_bytes()).hexdigest()
        if artifact_path.is_file()
        else hashlib.sha256(b"").hexdigest()
    )
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
