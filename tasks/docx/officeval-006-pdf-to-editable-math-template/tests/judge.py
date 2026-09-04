#!/usr/bin/env python3
"""officeval_006 的 Harbor Judge。

在上游格式评分之外，补充真实 PDF 的 21 道题、答案、OMML、图示、模板保留和
实际渲染检查，避免只贴五张页面截图也能通过。
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
from lxml import etree

import upstream_verifier


OUTPUT_NAME = "2026年3月17日高中数学作业_可编辑版.docx"
TEMPLATE_NAME = "高一年级第一次数学周练111 (2).docx"
TASK_ID = "cowork-evo/officeval-006-pdf-to-editable-math-template"
RUBRIC_VERSION = "omegause-officeval-006-adaptation-v0.1"
JUDGE_VERSION = "officeval-006-adaptation-v0.1"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W_NS, "m": M_NS}

CRITERION_SPECS = {
    "R001": ("指定 DOCX 可打开，主体正文与公式可编辑且不是页面截图", 3, "hurdle"),
    "R002": ("21 道真实试题、选项与小问按 PDF 顺序完整录入", 7, "positive"),
    "R003": ("1–17 题答案表及 1–21 题解答过程完整正确", 5, "positive"),
    "R004": ("分数和根式使用原生 OMML，没有普通字符替代", 4, "positive"),
    "R005": ("标题、章节、正文、选项和页码通过上游 11 分格式规则", 4, "positive"),
    "R006": ("抛物线图和等腰三角形图均保留", 2, "positive"),
    "R007": ("LibreOffice 实际渲染为五页且关键内容可读", 3, "positive"),
    "R008": ("模板页面几何及 PAGE/NUMPAGES 动态页脚保留", 2, "positive"),
    "R009": ("三道及以上试题缺失，或答案/解答大段缺失", -8, "penalty"),
    "R010": ("五张页面图片或不可编辑对象替代主体正文", -10, "penalty"),
    "R011": ("数学分数或根式大批量退化为普通文本", -6, "penalty"),
    "R012": ("渲染页数错误或关键内容、页码不可读", -5, "penalty"),
}


def normalized(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


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


def math_linear(element) -> str:
    """把单元格中的 OMML 线性化，仅用于答案表逐格核对。"""
    local = etree.QName(element).localname
    namespace = etree.QName(element).namespace
    if namespace in (W_NS, M_NS) and local == "t":
        return element.text or ""
    if namespace == M_NS and local == "f":
        num = element.find(f"{{{M_NS}}}num")
        den = element.find(f"{{{M_NS}}}den")
        return f"{math_linear(num)}/{math_linear(den)}"
    if namespace == M_NS and local == "rad":
        expr = element.find(f"{{{M_NS}}}e")
        return f"√{math_linear(expr)}"
    return "".join(math_linear(child) for child in element)


def table_rows(document: Document) -> list[list[list[str]]]:
    tables = []
    for table in document.tables:
        rows = []
        for row in table.rows:
            rows.append([normalized(math_linear(cell._tc)) for cell in row.cells])
        tables.append(rows)
    return tables


def answer_check(document: Document) -> tuple[bool, list[str]]:
    expected_first = ["C", "B", "A", "D", "C", "B", "A", "C", "B", "D"]
    expected_second = ["ACD", "ABC", "CD", "ABC", "[1,3)∪(3,+∞)", "4/5", "2"]
    found = []
    for rows in table_rows(document):
        if len(rows) >= 2 and rows[0] and rows[0][0] == "题号" and rows[1][0] == "答案":
            found.append(rows)
    issues = []
    if len(found) != 2:
        issues.append(f"答案表数量={len(found)}，期望=2")
    else:
        if found[0][1][1:] != expected_first:
            issues.append(f"1–10 答案={found[0][1][1:]}")
        if found[1][1][1:] != expected_second:
            issues.append(f"11–17 答案={found[1][1][1:]}")
    full_text = normalized("\n".join(p.text for p in document.paragraphs))
    if full_text.count("故选") < 14:
        issues.append(f"选择题解析中的‘故选’数量={full_text.count('故选')}，期望至少14")
    required_conclusions = [
        "定义域为[1,3)∪(3,+∞)",
        "a=2",
        "首项为2、公比为2的等比数列",
        "故MN=",
        "相切时判别式为0",
        "最大值为20，最小值为0",
    ]
    missing = [item for item in required_conclusions if normalized(item) not in full_text]
    if missing:
        issues.append(f"解答结论缺失={missing}")
    return not issues, issues


def render_check(path: Path) -> tuple[bool, str]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    pdfinfo = shutil.which("pdfinfo")
    if not libreoffice or not pdftotext or not pdfinfo:
        return False, "缺少 libreoffice、pdftotext 或 pdfinfo"
    with tempfile.TemporaryDirectory(prefix="cowork-evo-006-") as temp_dir:
        temp_root = Path(temp_dir)
        profile_uri = (temp_root / "lo-profile").resolve().as_uri()
        converted = subprocess.run(
            [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}",
             "--convert-to", "pdf", "--outdir", temp_dir, str(path)],
            capture_output=True, text=True, timeout=300, check=False,
        )
        pdf_path = temp_root / f"{path.stem}.pdf"
        if converted.returncode != 0 or not pdf_path.is_file():
            return False, (converted.stderr or converted.stdout or "PDF 转换失败")[-1000:]
        info = subprocess.run([pdfinfo, str(pdf_path)], capture_output=True, text=True,
                              timeout=120, check=False)
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        extracted = subprocess.run([pdftotext, "-layout", str(pdf_path), "-"],
                                   capture_output=True, text=True, timeout=120, check=False)
        rendered = normalized(extracted.stdout)
    required = [
        "2026年3月17日高中数学作业", "已知集合", "抛物线", "四、解答题",
        "参考答案", "定义域为", "最大值为20", "第1页共5页", "第5页共5页",
    ]
    missing = [item for item in required if normalized(item) not in rendered]
    return pages == 5 and not missing, f"页数={pages}，渲染缺失={missing}"


def footer_and_geometry_check(document: Document, template_path: Path) -> tuple[bool, list[str]]:
    issues = []
    template = Document(template_path)
    if len(document.sections) != len(template.sections):
        issues.append(f"section 数={len(document.sections)}，模板={len(template.sections)}")
    for index, (actual, source) in enumerate(zip(document.sections, template.sections), start=1):
        attrs = ("page_width", "page_height", "top_margin", "bottom_margin", "left_margin", "right_margin")
        for attr in attrs:
            if getattr(actual, attr) != getattr(source, attr):
                issues.append(f"第{index}节 {attr} 改变")
        xml = etree.tostring(actual.footer._element, encoding="unicode")
        if "PAGE" not in xml or "NUMPAGES" not in xml:
            issues.append(f"第{index}节页脚缺少 PAGE/NUMPAGES")
    return not issues, issues


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    expectations = json.loads(Path(__file__).with_name("expectations.json").read_text(encoding="utf-8"))
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)

    try:
        document = Document(output_path)
        with zipfile.ZipFile(output_path) as archive:
            xml_bytes = archive.read("word/document.xml")
            root = etree.fromstring(xml_bytes)
            media = [name for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/")]
            media_sizes = [len(archive.read(name)) for name in media]
    except Exception as exc:
        results.append(criterion("R001", False, f"DOCX 无法打开：{type(exc).__name__}: {exc}"))
        return build_result(results, output_path)

    regular_text = "".join(root.xpath(".//w:t/text()", namespaces=NS))
    all_text = "".join(root.xpath(".//w:t/text() | .//m:t/text()", namespaces=NS))
    editable_ok = len(normalized(regular_text)) >= 1200 and len(document.paragraphs) >= 45
    results.append(criterion("R001", editable_ok,
                             f"可编辑普通文字={len(normalized(regular_text))}字符，正文段落={len(document.paragraphs)}"))

    normalized_all = normalized(all_text)
    missing_questions = [fragment for fragment in expectations["question_fragments"]
                         if normalized(fragment) not in normalized_all]
    choice_tables = 0
    for rows in table_rows(document):
        if rows and rows[0] and rows[0][0].startswith("A．"):
            choice_tables += 1
    questions_ok = not missing_questions and choice_tables == 14
    results.append(criterion("R002", questions_ok,
                             f"缺失题目片段={missing_questions}，选择题选项表={choice_tables}/14"))

    answers_ok, answer_issues = answer_check(document)
    results.append(criterion("R003", answers_ok,
                             "两张答案表与21题解析完整" if answers_ok else "；".join(answer_issues)))

    fractions = len(root.xpath(".//m:f", namespaces=NS))
    radicals = len(root.xpath(".//m:rad", namespaces=NS))
    plain_text = "".join(root.xpath(".//w:t/text()", namespaces=NS))
    plain_sqrt = "√" in plain_text
    plain_fractions = re.findall(r"\d+\s*/\s*\d+", plain_text)
    formulas_ok = (fractions >= expectations["minimum_omml_fractions"]
                   and radicals >= expectations["minimum_omml_radicals"]
                   and not plain_sqrt and not plain_fractions)
    results.append(criterion("R004", formulas_ok,
                             f"OMML分数={fractions}，根式={radicals}，普通根号={plain_sqrt}，普通分数={plain_fractions[:8]}"))

    upstream = upstream_verifier.evaluate(str(artifact_dir))
    upstream_penalties = [item["rule"] for item in upstream.get("dim2_items", [])
                          if item.get("max_delta", 0) < 0 and item.get("hit")]
    upstream_ok = (upstream.get("dim1_pass") is True
                   and upstream.get("total_score") == 11
                   and upstream.get("max_score") == 11
                   and not upstream_penalties)
    results.append(criterion("R005", upstream_ok,
                             f"上游得分={upstream.get('total_score')}/{upstream.get('max_score')}，罚项={upstream_penalties}"))

    images_ok = 2 <= len(media) <= 4 and all(size >= 3000 for size in media_sizes[:2])
    results.append(criterion("R006", images_ok,
                             f"嵌入图片={len(media)}，字节数={media_sizes}"))

    rendered_ok, render_evidence = render_check(output_path)
    results.append(criterion("R007", rendered_ok, render_evidence))

    template_path = Path(__file__).resolve().parents[1] / "data" / "input_files" / TEMPLATE_NAME
    geometry_ok, geometry_issues = footer_and_geometry_check(document, template_path)
    results.append(criterion("R008", geometry_ok,
                             "页面几何与动态页脚均保留" if geometry_ok else "；".join(geometry_issues)))

    missing_answer_scale = 0 if answers_ok else len(answer_issues)
    results.append(criterion("R009", len(missing_questions) >= 3 or missing_answer_scale >= 3,
                             f"缺失题目={len(missing_questions)}，答案问题={missing_answer_scale}"))
    screenshot_substitution = len(media) >= 5 or len(normalized(regular_text)) < 1200
    results.append(criterion("R010", screenshot_substitution,
                             f"图片={len(media)}，可编辑普通文字={len(normalized(regular_text))}"))
    results.append(criterion("R011", not formulas_ok,
                             f"OMML分数={fractions}，根式={radicals}，普通分数={len(plain_fractions)}"))
    results.append(criterion("R012", not rendered_ok, render_evidence))
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
    hurdle_failures = [item["criterion_id"] for item in results
                       if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = 0.0
    digest = (hashlib.sha256(artifact_path.read_bytes()).hexdigest()
              if artifact_path.is_file() else hashlib.sha256(b"").hexdigest())
    normalized_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        normalized_results.append({
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
        })
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
