#!/usr/bin/env python3
"""FORTE Legal-020 的确定性 Harbor Judge。"""

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
from docx.document import Document as DocumentObject
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


OUTPUT_NAME = "授权委托书风险提示报告.docx"
TASK_ID = "cowork-evo/forte-legal-020-power-of-attorney-risk-report"
RUBRIC_VERSION = "forte-legal-020-adaptation-v0.1"
JUDGE_VERSION = "forte-legal-020-adaptation-v0.2"

EXPECTED_LEVELS = {
    1: "中风险",
    2: "中风险",
    3: "中风险",
    4: "高风险",
    5: "高风险",
    6: "中风险",
}

EXPECTED_RISKS = {
    1: {"受托人无相应资质", "转委托约定缺失", "法律责任承担约定缺失"},
    2: {"特别授权范围过宽", "转委托约定缺失", "法律责任承担约定缺失"},
    3: {"受托人无相应资质", "转委托约定缺失", "法律责任承担约定缺失"},
    4: {"委托人身份证明缺失", "受托人身份证明缺失", "转委托约定缺失", "法律责任承担约定缺失"},
    5: {"授权范围完全未明确", "受托人无相应资质", "转委托约定缺失", "法律责任承担约定缺失"},
    6: {"特别授权范围过宽", "转委托约定缺失", "法律责任承担约定缺失"},
}

SOURCE_ANCHORS = {
    1: {"张伟", "李芳", "（2026）京0101民初第0012号"},
    2: {"陈雪梅", "王建国律师", "（2026）浙0101民初第0588号"},
    3: {"刘洋", "赵明辉", "（2023）粤0305民初第1234号"},
    4: {"周丽华", "孙志强律师", "（2025）京0115民初第0336号"},
    5: {"吴强", "黄丽萍", "（2026）川0107民初第0215号"},
    6: {"上海盛宏贸易有限公司", "陈明律师", "（2026）苏0101民初第0412号"},
}

KNOWN_RISK_LABELS = {
    "委托人身份证明缺失",
    "受托人身份证明缺失",
    "委托人与业务主体不匹配",
    "受托人无相应资质",
    "案件信息不完整",
    "授权范围完全未明确",
    "特别授权范围过宽",
    "授权事项笼统无边界",
    "授权事项与业务场景不匹配",
    "授权期限完全缺失",
    "授权期限过长或无限期",
    "授权起止日期不完整",
    "委托日期缺失",
    "转委托约定缺失",
    "转委托限制条款不完整",
    "法律责任承担约定缺失",
    "免责条款表述存在瑕疵",
    "签字/盖章完全缺失",
    "受托人姓名或名称缺失",
    "企业委托书仅有签字无公章",
    "附件关联缺失",
}

CRITERION_SPECS = {
    "R001": ("指定 DOCX 可打开、可编辑并包含完整报告结构", 3, "hurdle"),
    "R002": ("六份委托书均有独立节和五项必要内容", 5, "hurdle"),
    "R003": ("委托书1和3的等级与风险集合正确", 4, "hurdle"),
    "R004": ("委托书2和6的等级与风险集合正确", 4, "hurdle"),
    "R005": ("委托书4的等级与风险集合正确", 4, "hurdle"),
    "R006": ("委托书5的等级与风险集合正确", 4, "hurdle"),
    "R007": ("未签署状态已排除且没有额外风险项", 3, "hurdle"),
    "R008": ("原文证据、风险说明和调整意见完整可追溯", 4, "hurdle"),
    "R009": ("报告实际渲染清晰且无空白页", 3, "positive"),
    "R010": ("两份及以上文件的等级或真值集合错误", -9, "penalty"),
    "R011": ("错误触发未签署风险或累计两个以上额外风险", -7, "penalty"),
    "R012": ("正文截图化、明显截断或无法渲染", -10, "penalty"),
}

HURDLE_CAPS = {"R001": 0.0, **{f"R{number:03d}": 0.74 for number in range(2, 9)}}


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


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def iter_block_texts(document: DocumentObject) -> list[str]:
    """按正文顺序抽取段落和表格行，保留节间边界。"""

    blocks: list[str] = []
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            text = Paragraph(child, document).text.strip()
            if text:
                blocks.append(text)
        elif isinstance(child, CT_Tbl):
            table = Table(child, document)
            for row in table.rows:
                values = []
                for cell in row.cells:
                    value = " ".join(paragraph.text.strip() for paragraph in cell.paragraphs if paragraph.text.strip())
                    if not values or value != values[-1]:
                        values.append(value)
                text = " | ".join(values).strip()
                if text:
                    blocks.append(text)
    return blocks


def extract_detail_sections(blocks: list[str]) -> tuple[dict[int, str], list[str]]:
    candidates: dict[int, list[int]] = {number: [] for number in range(1, 7)}
    for index, block in enumerate(blocks):
        match = re.match(r"^委托书\s*([1-6])(?:\s|[：:｜|]|$)", block)
        if match:
            candidates[int(match.group(1))].append(index)
    issues = [f"委托书{number}缺少独立标题" for number, positions in candidates.items() if not positions]
    if issues:
        return {}, issues

    # 总览表也可能以“委托书N”开头；每份最后一次出现才是详细核查标题。
    starts = {number: positions[-1] for number, positions in candidates.items()}
    ordered = sorted((position, number) for number, position in starts.items())
    if [number for _, number in ordered] != list(range(1, 7)):
        issues.append(f"详细章节顺序异常：{[number for _, number in ordered]}")
    sections = {}
    for offset, (position, number) in enumerate(ordered):
        end = ordered[offset + 1][0] if offset + 1 < len(ordered) else len(blocks)
        sections[number] = "\n".join(blocks[position:end])
    return sections, issues


def label_is_triggered(section: str, label: str) -> bool:
    """忽略明确写成“排除/未触发”的标签提及。"""

    for match in re.finditer(re.escape(label), section):
        context = section[max(0, match.start() - 45) : min(len(section), match.end() + 45)]
        if any(marker in context for marker in ("排除", "未触发", "不触发", "不计入", "未检出", "未纳入")):
            continue
        return True
    return False


def observed_risks(section: str) -> set[str]:
    return {label for label in KNOWN_RISK_LABELS if label_is_triggered(section, label)}


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
    with tempfile.TemporaryDirectory(prefix="cowork-evo-legal020-") as temp_dir:
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
        rendered = compact("\n".join(pages))
        substantive_lengths = []
        for page in pages:
            cleaned = re.sub(r"FORTE\s*Legal-020.*?授权委托书风险提示报告", "", page, flags=re.I)
            cleaned = re.sub(r"第\s*\d+\s*页", "", cleaned)
            substantive_lengths.append(len(compact(cleaned)))

    required = [
        "授权委托书风险提示报告",
        *[f"委托书{number}" for number in range(1, 7)],
        "受托人无相应资质",
        "特别授权范围过宽",
        "委托人身份证明缺失",
        "授权范围完全未明确",
        "转委托约定缺失",
        "法律责任承担约定缺失",
    ]
    missing = [item for item in required if compact(item) not in rendered]
    empty_pages = [index + 1 for index, length in enumerate(substantive_lengths) if length < 80]
    # 中文报告去除空白后字符密度远高于英文；完整 Golden 为 4236 字符。
    # 3500 的下限允许简洁表达，但会拒绝只有结论清单、没有逐项说明的报告。
    ok = 4 <= len(pages) <= 12 and not missing and not empty_pages and len(rendered) >= 3500
    evidence = f"页数={len(pages)}，渲染缺失={missing}，近空白页={empty_pages}，逐页字符={substantive_lengths}"
    return ok, evidence, len(pages)


def check_pair(numbers: tuple[int, ...], levels: dict[int, bool], sets: dict[int, bool]) -> tuple[bool, str]:
    details = []
    for number in numbers:
        details.append(f"{number}:等级={levels.get(number, False)},集合={sets.get(number, False)}")
    return all(levels.get(number, False) and sets.get(number, False) for number in numbers), "；".join(details)


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)

    try:
        document = Document(output_path)
        blocks = iter_block_texts(document)
        full_text = "\n".join(blocks)
        media_count = count_media(output_path)
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        results.append(criterion("R001", False, f"DOCX 无法读取：{error}"))
        return build_result(results, output_path)

    sections, section_issues = extract_detail_sections(blocks)
    editable_ok = len(compact(full_text)) >= 3500 and len(document.tables) >= 6 and media_count <= 4
    results.append(
        criterion(
            "R001",
            editable_ok and len(sections) == 6,
            f"正文字符={len(compact(full_text))}，表格={len(document.tables)}，媒体={media_count}，章节问题={section_issues}",
        )
    )

    component_checks = {}
    for number, section in sections.items():
        component_checks[number] = all(
            item in section
            for item in (EXPECTED_LEVELS[number], "风险项", "触发证据", "风险说明", "调整意见")
        ) and len(compact(section)) >= 450
    all_components = len(component_checks) == 6 and all(component_checks.values()) and not section_issues
    results.append(criterion("R002", all_components, f"逐份组件={component_checks}，章节问题={section_issues}"))

    observed = {number: observed_risks(section) for number, section in sections.items()}
    risk_set_checks = {
        number: observed.get(number, set()) == EXPECTED_RISKS[number]
        for number in range(1, 7)
    }
    level_checks = {
        number: EXPECTED_LEVELS[number] in sections.get(number, "")
        for number in range(1, 7)
    }

    ok, evidence = check_pair((1, 3), level_checks, risk_set_checks)
    results.append(criterion("R003", ok, evidence + f"；实际集合={{1:{sorted(observed.get(1, set()))},3:{sorted(observed.get(3, set()))}}}"))
    ok, evidence = check_pair((2, 6), level_checks, risk_set_checks)
    results.append(criterion("R004", ok, evidence + f"；实际集合={{2:{sorted(observed.get(2, set()))},6:{sorted(observed.get(6, set()))}}}"))
    ok, evidence = check_pair((4,), level_checks, risk_set_checks)
    results.append(criterion("R005", ok, evidence + f"；实际集合={sorted(observed.get(4, set()))}"))
    ok, evidence = check_pair((5,), level_checks, risk_set_checks)
    results.append(criterion("R006", ok, evidence + f"；实际集合={sorted(observed.get(5, set()))}"))

    improper_signature = any(
        label_is_triggered(section, "签字/盖章完全缺失")
        for section in sections.values()
    )
    unexpected = {
        number: observed.get(number, set()) - EXPECTED_RISKS[number]
        for number in range(1, 7)
    }
    exact_sets = all(risk_set_checks.values())
    results.append(
        criterion(
            "R007",
            exact_sets and not improper_signature,
            f"精确集合={risk_set_checks}，未签署误触发={improper_signature}，额外项={unexpected}",
        )
    )

    source_checks = {}
    action_checks = {}
    for number in range(1, 7):
        section = sections.get(number, "")
        section_compact = compact(section)
        source_checks[number] = all(compact(anchor) in section_compact for anchor in SOURCE_ANCHORS[number])
        action_checks[number] = (
            "触发证据" in section
            and "风险说明" in section
            and "调整意见" in section
            and any(verb in section for verb in ("补充", "明确", "增加", "改为", "更换", "收窄"))
            and len(compact(section)) >= 450
        )
    evidence_ok = all(source_checks.values()) and all(action_checks.values())
    results.append(criterion("R008", evidence_ok, f"来源锚点={source_checks}，说明与建议={action_checks}"))

    render_ok, render_evidence, _ = render_check(output_path)
    results.append(criterion("R009", render_ok, render_evidence))

    mismatch_documents = [
        number
        for number in range(1, 7)
        if not level_checks[number] or not risk_set_checks[number]
    ]
    results.append(criterion("R010", len(mismatch_documents) >= 2, f"错误文件={mismatch_documents}"))

    unexpected_count = sum(len(values) for values in unexpected.values())
    results.append(
        criterion(
            "R011",
            improper_signature or unexpected_count >= 2,
            f"未签署误触发={improper_signature}，额外风险累计={unexpected_count}，明细={unexpected}",
        )
    )

    rasterized_or_broken = media_count > 0 and len(compact(full_text)) < 3000
    unusable = rasterized_or_broken or len(compact(full_text)) < 2500 or not render_ok
    results.append(
        criterion(
            "R012",
            unusable,
            f"媒体={media_count}，正文字符={len(compact(full_text))}，渲染通过={render_ok}",
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
