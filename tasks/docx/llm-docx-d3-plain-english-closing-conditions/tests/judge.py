#!/usr/bin/env python3
"""LLM-DOCX-Editing D3 的确定性 Harbor Judge。"""

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


OUTPUT_NAME = "cap-conditional-edit-plain-english.docx"
INPUT_NAME = "cap-conditional-edit.docx"
TASK_ID = "cowork-evo/llm-docx-d3-plain-english-closing-conditions"
RUBRIC_VERSION = "llm-docx-editing-d3-adaptation-v0.1"
JUDGE_VERSION = "llm-docx-editing-d3-adaptation-v0.1"
HURDLE_CAPS = {"R001": 0.0, "R002": 0.74, "R007": 0.74}

HEADINGS = [
    "Parties",
    "Transaction Structure",
    "Purchase Price",
    "Representations and Warranties",
    "Indemnification",
    "Closing Conditions",
    "Governing Law",
    "Exclusivity",
]

CRITERION_SPECS = {
    "R001": ("指定 DOCX 可打开、正文可编辑且 term sheet 完整", 3, "hurdle"),
    "R002": ("五项成交条件均以可理解英语完整保留", 6, "hurdle"),
    "R003": ("HSR clearance 已解释为政府反垄断批准", 3, "positive"),
    "R004": ("material adverse change 已解释为无重大负面变化", 3, "positive"),
    "R005": ("第三方批准和买方融资条件准确", 3, "positive"),
    "R006": ("目标段落是完整、面向非法律读者的英语", 3, "positive"),
    "R007": ("目标节以外全部章节和事实原样保持", 5, "hurdle"),
    "R008": ("八个章节顺序与可编辑结构保持", 3, "positive"),
    "R009": ("实际渲染后金额、章节和改写均可读", 3, "positive"),
    "R010": ("五项条件中缺少至少一项", -8, "penalty"),
    "R011": ("修改或删除了目标节以外内容", -8, "penalty"),
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


def section_bounds(paragraphs: list[str], heading: str, next_heading: str) -> tuple[int, int] | None:
    try:
        start = paragraphs.index(heading)
        end = paragraphs.index(next_heading, start + 1)
    except ValueError:
        return None
    return start, end


def condition_checks(target: str) -> dict[str, bool]:
    text = normalized(target)
    return {
        "a_representations_accurate": (
            any(term in text for term in ("promises", "statements", "stated", "representations"))
            and "true" in text
            and "accurate" in text
            and any(term in text for term in ("closing", "sale", "completed"))
        ),
        "b_antitrust_approval": "government" in text and "antitrust" in text and "approval" in text,
        "c_no_negative_change": (
            "no significant negative change" in text
            and "business" in text
        ),
        "d_third_party_schedule_3": (
            "third-party" in text
            and any(term in text for term in ("approval", "permission", "consent"))
            and "schedule 3" in text
        ),
        "e_buyer_financing": (
            "buyer" in text
            and "financ" in text
            and any(term in text for term in ("secured", "obtained", "completed", "available"))
        ),
    }


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
    with tempfile.TemporaryDirectory(prefix="cowork-evo-d3-") as temp_dir:
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
        "$15,000,000",
        "$12,000,000",
        "$2,000,000",
        "$1,000,000",
        "$250,000",
        "$7,500,000",
        "government antitrust approval",
        "no significant negative changes to the business",
        "Schedule 3",
        "Buyer",
        "financing",
    ]
    missing = [item for item in required if normalized(item) not in rendered]
    ok = 1 <= len(pages) <= 2 and not missing and len(rendered) >= 1900
    return ok, f"页数={len(pages)}，渲染缺失={missing}，文本字符={len(rendered)}", len(pages)


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
        media_count = count_media(output_path)
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        results.append(criterion("R001", False, f"DOCX 无法读取：{error}"))
        return build_result(results, output_path)

    bounds = section_bounds(paragraphs, "Closing Conditions", "Governing Law")
    editable_ok = len(paragraphs) >= 20 and len(normalized(full_text)) >= 1900 and media_count == 0 and bounds is not None
    results.append(
        criterion(
            "R001",
            editable_ok,
            f"非空段落={len(paragraphs)}，文本字符={len(normalized(full_text))}，媒体={media_count}，目标节={bounds}",
        )
    )

    if bounds is None:
        target = ""
        close_index = -1
        governing_index = -1
    else:
        close_index, governing_index = bounds
        target = "\n".join(paragraphs[close_index + 1 : governing_index])
    conditions = condition_checks(target)
    all_conditions = all(conditions.values())
    results.append(criterion("R002", all_conditions, f"五项条件={conditions}"))

    target_norm = normalized(target)
    hsr_explained = conditions["b_antitrust_approval"] and "hsr clearance" not in target_norm
    results.append(criterion("R003", hsr_explained, f"反垄断解释={conditions['b_antitrust_approval']}，残留 HSR={'hsr clearance' in target_norm}"))

    mac_explained = conditions["c_no_negative_change"] and "material adverse change" not in target_norm
    results.append(criterion("R004", mac_explained, f"负面变化解释={conditions['c_no_negative_change']}，残留 MAC={'material adverse change' in target_norm}"))

    de_ok = conditions["d_third_party_schedule_3"] and conditions["e_buyer_financing"]
    results.append(criterion("R005", de_ok, f"第三方批准={conditions['d_third_party_schedule_3']}，买方融资={conditions['e_buyer_financing']}"))

    words = re.findall(r"[A-Za-z0-9$'-]+", target)
    plain_opening = any(phrase in target_norm for phrase in ("before the sale", "sale can only", "before closing", "to complete the sale"))
    jargon_absent = all(term not in target_norm for term in ("hsr clearance", "material adverse change", "hereunder", "thereof", "pursuant to"))
    plain_ok = 45 <= len(words) <= 180 and plain_opening and jargon_absent
    results.append(criterion("R006", plain_ok, f"单词数={len(words)}，直白开头={plain_opening}，术语清理={jargon_absent}"))

    source_path = Path(__file__).resolve().parents[1] / "data" / "input_files" / INPUT_NAME
    source_document = Document(source_path)
    source_paragraphs = [paragraph.text.strip() for paragraph in source_document.paragraphs if paragraph.text.strip()]
    source_bounds = section_bounds(source_paragraphs, "Closing Conditions", "Governing Law")
    if bounds is None or source_bounds is None:
        prefix_ok = suffix_ok = False
        outside_evidence = "目标或来源章节边界缺失"
    else:
        source_close, source_governing = source_bounds
        prefix_expected = source_paragraphs[: source_close + 1]
        suffix_expected = source_paragraphs[source_governing:]
        prefix_actual = paragraphs[: close_index + 1]
        suffix_actual = paragraphs[governing_index:]
        prefix_ok = prefix_actual == prefix_expected
        suffix_ok = suffix_actual == suffix_expected
        outside_evidence = (
            f"目标节前一致={prefix_ok}({len(prefix_actual)}/{len(prefix_expected)})，"
            f"目标节后一致={suffix_ok}({len(suffix_actual)}/{len(suffix_expected)})"
        )
    preservation_ok = prefix_ok and suffix_ok
    results.append(criterion("R007", preservation_ok, outside_evidence))

    heading_positions = []
    for heading in HEADINGS:
        try:
            heading_positions.append(paragraphs.index(heading))
        except ValueError:
            heading_positions.append(-1)
    structure_ok = all(position >= 0 for position in heading_positions) and heading_positions == sorted(heading_positions) and len(target_norm) >= 350
    results.append(criterion("R008", structure_ok, f"章节位置={dict(zip(HEADINGS, heading_positions))}，目标字符={len(target_norm)}"))

    render_ok, render_evidence, _ = render_check(output_path)
    results.append(criterion("R009", render_ok, render_evidence))

    missing_conditions = [name for name, passed in conditions.items() if not passed]
    results.append(criterion("R010", bool(missing_conditions), f"缺失条件={missing_conditions}"))
    results.append(criterion("R011", not preservation_ok, outside_evidence))

    rasterized_or_broken = media_count > 0 and len(normalized(full_text)) < 1500
    unusable = rasterized_or_broken or len(normalized(full_text)) < 1200 or not render_ok
    results.append(criterion("R012", unusable, f"媒体={media_count}，正文字符={len(normalized(full_text))}，渲染通过={render_ok}"))
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
