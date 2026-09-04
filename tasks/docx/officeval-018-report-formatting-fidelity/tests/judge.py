#!/usr/bin/env python3
"""officeval_018 的 Harbor Judge。

上游 Verifier 原样保存在同目录，但它把输入中不存在的“临海”标题硬编码为
必选文字，并且三/四级标题候选正则互相重叠。本 Judge 以真实输入和版式
要求为准，对这些不可同时满足的条件做了逐题修正。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


TASK_ID = "cowork-evo/officeval-018-report-formatting-fidelity"
RUBRIC_VERSION = "omegause-officeval-018-adaptation-v0.1"
JUDGE_VERSION = "officeval-018-adaptation-v0.1"
PASS_THRESHOLD = 0.75

BASE_DIR = Path(__file__).resolve().parents[1]
EXPECTATIONS = json.loads((Path(__file__).with_name("expectations.json")).read_text(encoding="utf-8"))
OUTPUT_NAME = EXPECTATIONS["output_filename"]
INPUT_NAME = EXPECTATIONS["input_filename"]
TITLE = EXPECTATIONS["title"]

CRITERION_SPECS = {
    "R001": ("指定名称的 DOCX 可正常打开", 2, "hurdle"),
    "R002": ("原始正文、表格数据、18 个原生图表和可见绘图对象完整保留", 7, "hurdle"),
    "R003": ("题目与摘要页的内容、顺序和段落格式正确", 4, "positive"),
    "R004": ("一至五级标题使用正确层次和编号系统", 6, "positive"),
    "R005": ("题目、摘要与各级标题的字体、字号和字重正确", 5, "positive"),
    "R006": ("正文为宋体小四，1.25 倍行距、段前 0.5 行、首行缩进两字符", 5, "positive"),
    "R007": ("31 组可见图形的图注位置、连续编号、内容与格式正确", 6, "positive"),
    "R008": ("25 张表的表注位置、连续编号、内容与格式正确", 6, "positive"),
    "R009": ("表内文字为宋体小四单倍行距，长表跨页重复表头且不拆单行", 4, "positive"),
    "R010": ("所有节均使用上下 2.54 cm、左右 3.18 cm 页边距", 3, "positive"),
    "R011": ("文档可由 LibreOffice 实际渲染，页数与关键内容合理", 4, "positive"),
    "R012": ("缺失 5 条及以上应有的图注或表注", -6, "penalty"),
    "R013": ("大量丢失正文、表格数据或原生图表", -8, "penalty"),
    "R014": ("多个图表题注位置错乱或标题编号叠加", -4, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
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


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def direct_rpr_values(rpr) -> tuple[str | None, float | None, bool | None]:
    if rpr is None:
        return None, None, None
    name = None
    size = None
    bold = None
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is not None:
        name = fonts.get(qn("w:eastAsia")) or fonts.get(qn("w:ascii"))
    size_element = rpr.find(qn("w:sz"))
    if size_element is not None:
        try:
            size = float(size_element.get(qn("w:val"))) / 2.0
        except (TypeError, ValueError):
            size = None
    bold_element = rpr.find(qn("w:b"))
    if bold_element is not None:
        bold = bold_element.get(qn("w:val")) not in ("0", "false")
    return name, size, bold


def effective_run_values(document: Document, paragraph, run) -> tuple[str | None, float | None, bool]:
    layers = [run._element.rPr]
    try:
        style = run.style
    except Exception:
        style = None
    while style is not None:
        layers.append(style.element.find(qn("w:rPr")))
        style = style.base_style
    ppr = paragraph._element.find(qn("w:pPr"))
    if ppr is not None:
        layers.append(ppr.find(qn("w:rPr")))
    try:
        style = paragraph.style
    except Exception:
        style = None
    while style is not None:
        layers.append(style.element.find(qn("w:rPr")))
        style = style.base_style
    defaults = document.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        default_run = defaults.find(qn("w:rPrDefault"))
        if default_run is not None:
            layers.append(default_run.find(qn("w:rPr")))

    name = size = bold = None
    for layer in layers:
        layer_name, layer_size, layer_bold = direct_rpr_values(layer)
        if name is None and layer_name is not None:
            name = layer_name
        if size is None and layer_size is not None:
            size = layer_size
        if bold is None and layer_bold is not None:
            bold = layer_bold
    return name, size, bool(bold)


def font_matches(document: Document, paragraph, name: str, size: float, bold: bool | None = None) -> bool:
    checked = 0
    accepted_names = {name}
    if name == "宋体":
        accepted_names.add("SimSun")
    if name == "黑体":
        accepted_names.add("SimHei")
    for run in paragraph.runs:
        if not run.text.strip():
            continue
        checked += 1
        actual_name, actual_size, actual_bold = effective_run_values(document, paragraph, run)
        if actual_name not in accepted_names or actual_size is None or abs(actual_size - size) > 0.5:
            return False
        if bold is not None and actual_bold is not bold:
            return False
    return checked > 0


def paragraph_spacing_matches(paragraph, *, line: int, before_lines: int, first_line_chars: int) -> bool:
    ppr = paragraph._element.find(qn("w:pPr"))
    if ppr is None:
        return False
    spacing = ppr.find(qn("w:spacing"))
    indent = ppr.find(qn("w:ind"))
    if spacing is None or indent is None:
        return False
    return (
        spacing.get(qn("w:line")) == str(line)
        and spacing.get(qn("w:lineRule")) in (None, "auto")
        and spacing.get(qn("w:beforeLines"), "0") == str(before_lines)
        and indent.get(qn("w:firstLineChars"), "0") == str(first_line_chars)
    )


def table_cells(table) -> list[str]:
    return [normalize_text(cell.text) for row in table.rows for cell in row.cells]


def document_content_fidelity(source: Document, candidate: Document) -> tuple[bool, dict]:
    source_blocks = [
        normalize_text(paragraph.text)
        for paragraph in source.paragraphs
        if len(normalize_text(paragraph.text)) >= 80
        and not re.match(r"^[图表]\s*\d+", paragraph.text.strip())
    ]
    candidate_text = normalize_text("\n".join(paragraph.text for paragraph in candidate.paragraphs))
    block_hits = sum(block in candidate_text for block in source_blocks)
    block_ratio = block_hits / len(source_blocks) if source_blocks else 0.0

    equal_tables = 0
    for source_table, candidate_table in zip(source.tables, candidate.tables):
        if table_cells(source_table) == table_cells(candidate_table):
            equal_tables += 1

    drawings = len(candidate.element.findall(".//" + qn("w:drawing")))
    drawing_paragraphs = sum(
        bool(paragraph._element.findall(".//" + qn("w:drawing")))
        for paragraph in candidate.paragraphs
    )
    charts = len(candidate.element.findall(".//" + qn("c:chart")))
    expected = EXPECTATIONS["object_counts"]
    # 上游输入含 1 个全透明 PNG 占位对象，它既可保留也可删除；
    # 两种处理都不应影响“可见内容保留”评分。
    drawing_count_ok = expected["drawing_objects"] <= drawings <= expected["drawing_objects"] + 1
    passed = (
        block_ratio >= EXPECTATIONS["body_block_min_ratio"]
        and equal_tables == expected["tables"]
        and len(candidate.tables) == expected["tables"]
        and drawing_count_ok
        and drawing_paragraphs == expected["drawing_paragraphs"]
        and charts == expected["charts"]
    )
    evidence = {
        "body_block_ratio": round(block_ratio, 6),
        "equal_tables": equal_tables,
        "tables": len(candidate.tables),
        "drawing_objects": drawings,
        "drawing_objects_accepted": drawing_count_ok,
        "drawing_paragraphs": drawing_paragraphs,
        "charts": charts,
    }
    return passed, evidence


def title_and_abstract_check(document: Document) -> tuple[bool, str]:
    nonempty = [(index, paragraph) for index, paragraph in enumerate(document.paragraphs) if paragraph.text.strip()]
    if len(nonempty) < 2:
        return False, "非空段落不足 2 个"
    title_index, title = nonempty[0]
    abstract_index, abstract = nonempty[1]
    title_ppr = title._element.find(qn("w:pPr"))
    title_spacing = title_ppr.find(qn("w:spacing")) if title_ppr is not None else None
    single_line = title_spacing is not None and title_spacing.get(qn("w:line")) in ("240", None)
    abstract_after_page_break = bool(abstract.paragraph_format.page_break_before)
    passed = (
        title.text.strip() == TITLE
        and abstract.text.strip() == "摘要"
        and title_index < abstract_index
        and title.alignment == WD_ALIGN_PARAGRAPH.CENTER
        and abstract.alignment == WD_ALIGN_PARAGRAPH.CENTER
        and single_line
        and abstract_after_page_break
    )
    return passed, (
        f"题目={title.text.strip()!r}，摘要={abstract.text.strip()!r}，"
        f"居中={title.alignment == WD_ALIGN_PARAGRAPH.CENTER}/{abstract.alignment == WD_ALIGN_PARAGRAPH.CENTER}，"
        f"摘要分页={abstract_after_page_break}"
    )


def heading_check(document: Document) -> tuple[bool, str, dict[int, list]]:
    patterns = {
        1: re.compile(r"^[一二三四五六七八九十]+、\S"),
        2: re.compile(r"^（[一二三四五六七八九十]+）\S"),
        3: re.compile(r"^\d+\.\S"),
        4: re.compile(r"^（\d+）\S"),
        5: re.compile(r"^\d+）\S"),
    }
    grouped: dict[int, list] = {}
    details = []
    passed = True
    for level in range(1, 6):
        paragraphs = [
            paragraph for paragraph in document.paragraphs
            if paragraph.style is not None and paragraph.style.name == f"Heading {level}" and paragraph.text.strip()
        ]
        grouped[level] = paragraphs
        expected_count = EXPECTATIONS["heading_counts"][str(level)]
        format_ok = all(patterns[level].match(paragraph.text.strip()) for paragraph in paragraphs)
        count_ok = len(paragraphs) == expected_count
        if level == 1:
            expected_first = [f"{value}、" for value in "一二三四五六七"]
            first_ok = [paragraph.text.strip()[:2] for paragraph in paragraphs] == expected_first
        else:
            first_ok = True
        level_ok = count_ok and format_ok and first_ok
        passed = passed and level_ok
        details.append(f"L{level}={len(paragraphs)}/{expected_count},format={format_ok},sequence={first_ok}")
    return passed, "；".join(details), grouped


def heading_font_check(document: Document, grouped: dict[int, list]) -> tuple[bool, str]:
    nonempty = [paragraph for paragraph in document.paragraphs if paragraph.text.strip()]
    title = nonempty[0] if nonempty else None
    abstract = nonempty[1] if len(nonempty) > 1 else None
    title_ok = title is not None and font_matches(document, title, "宋体", 16, True)
    abstract_ok = abstract is not None and font_matches(document, abstract, "黑体", 14, True)
    level_specs = {1: ("黑体", 15, True), 2: ("黑体", 14, True), 3: ("宋体", 12, False), 4: ("宋体", 12, False), 5: ("宋体", 12, False)}
    checked = correct = 0
    for level, paragraphs in grouped.items():
        for paragraph in paragraphs:
            checked += 1
            if font_matches(document, paragraph, *level_specs[level]):
                correct += 1
    ratio = correct / checked if checked else 0.0
    return title_ok and abstract_ok and ratio >= 0.98, f"题目={title_ok}，摘要={abstract_ok}，标题字体={correct}/{checked}"


def body_format_check(document: Document) -> tuple[bool, str]:
    targets = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text or text in (TITLE, "摘要"):
            continue
        if paragraph.style is not None and paragraph.style.name.startswith("Heading "):
            continue
        if re.match(r"^[图表]\s*\d+", text):
            continue
        if paragraph._element.findall(".//" + qn("w:drawing")):
            continue
        targets.append(paragraph)
    font_ok = sum(font_matches(document, paragraph, "宋体", 12) for paragraph in targets)
    spacing_ok = sum(
        paragraph_spacing_matches(paragraph, line=300, before_lines=50, first_line_chars=200)
        for paragraph in targets
    )
    count = len(targets)
    passed = count >= 150 and font_ok / count >= 0.98 and spacing_ok / count >= 0.98
    return passed, f"正文段={count}，字体={font_ok}/{count}，段落格式={spacing_ok}/{count}"


def child_text(child) -> str:
    return "".join(node.text or "" for node in child.iter() if node.tag == qn("w:t")).strip()


def next_nonempty_paragraph(children: list, start: int):
    for child in children[start + 1:]:
        if child.tag == qn("w:tbl"):
            return None
        if child.tag == qn("w:p") and child_text(child):
            return child
    return None


def previous_nonempty_paragraph(children: list, start: int):
    for child in reversed(children[:start]):
        if child.tag == qn("w:tbl"):
            return None
        if child.tag == qn("w:p") and child_text(child):
            return child
    return None


def top_level_paragraph_map(document: Document) -> dict[int, object]:
    return {id(paragraph._element): paragraph for paragraph in document.paragraphs}


def caption_check(document: Document, kind: str) -> tuple[bool, dict]:
    is_figure = kind == "figure"
    prefix = "图" if is_figure else "表"
    expected_count = EXPECTATIONS["object_counts"]["figure_captions" if is_figure else "table_captions"]
    keywords = EXPECTATIONS["figure_title_keywords" if is_figure else "table_title_keywords"]
    pattern = re.compile(rf"^{prefix} (\d+)  (\S.+)$")
    captions = []
    for paragraph in document.paragraphs:
        match = pattern.match(paragraph.text.strip())
        if match:
            captions.append((paragraph, int(match.group(1)), match.group(2).strip()))

    numbers = [item[1] for item in captions]
    titles = [item[2] for item in captions]
    sequence_ok = numbers == list(range(1, len(numbers) + 1))
    unique_ok = len(set(titles)) == len(titles) and all(len(title) >= 4 for title in titles)
    keyword_hits = 0
    for title, required in zip(titles, keywords):
        if all(token in title for token in required):
            keyword_hits += 1
    keyword_ok = len(captions) == expected_count and keyword_hits / expected_count >= 0.90
    format_ok = all(
        paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
        and font_matches(document, paragraph, "宋体", 12, False)
        for paragraph, _, _ in captions
    )

    children = list(document.element.body)
    paragraph_map = top_level_paragraph_map(document)
    positioned = 0
    targets = 0
    if is_figure:
        for index, child in enumerate(children):
            if child.tag != qn("w:p") or not child.findall(".//" + qn("w:drawing")):
                continue
            targets += 1
            adjacent = next_nonempty_paragraph(children, index)
            if adjacent is not None and pattern.match(child_text(adjacent)):
                positioned += 1
    else:
        for index, child in enumerate(children):
            if child.tag != qn("w:tbl"):
                continue
            targets += 1
            adjacent = previous_nonempty_paragraph(children, index)
            if adjacent is not None and pattern.match(child_text(adjacent)):
                positioned += 1
    position_ok = targets == expected_count and positioned == targets
    passed = (
        len(captions) == expected_count
        and sequence_ok
        and unique_ok
        and keyword_ok
        and format_ok
        and position_ok
    )
    return passed, {
        "captions": len(captions),
        "targets": targets,
        "positioned": positioned,
        "sequence": sequence_ok,
        "unique": unique_ok,
        "keyword_hits": keyword_hits,
        "format": format_ok,
    }


def table_format_check(document: Document) -> tuple[bool, str]:
    tables = document.tables
    text_runs = font_runs = paragraphs = spacing_paragraphs = 0
    cant_split_rows = header_tables = 0
    total_rows = 0
    for table in tables:
        first_tr_pr = table.rows[0]._tr.get_or_add_trPr() if table.rows else None
        if first_tr_pr is not None and first_tr_pr.find(qn("w:tblHeader")) is not None:
            header_tables += 1
        for row in table.rows:
            total_rows += 1
            tr_pr = row._tr.get_or_add_trPr()
            if tr_pr.find(qn("w:cantSplit")) is not None:
                cant_split_rows += 1
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if not paragraph.text.strip():
                        continue
                    paragraphs += 1
                    if paragraph_spacing_matches(paragraph, line=240, before_lines=0, first_line_chars=0):
                        spacing_paragraphs += 1
                    for run in paragraph.runs:
                        if not run.text.strip():
                            continue
                        text_runs += 1
                        actual_name, actual_size, _ = effective_run_values(document, paragraph, run)
                        if actual_name in ("宋体", "SimSun") and actual_size is not None and abs(actual_size - 12) <= 0.5:
                            font_runs += 1
    passed = (
        len(tables) == EXPECTATIONS["object_counts"]["tables"]
        and text_runs > 0
        and font_runs == text_runs
        and spacing_paragraphs == paragraphs
        and cant_split_rows == total_rows
        and header_tables == len(tables)
    )
    return passed, (
        f"表格={len(tables)}，字体run={font_runs}/{text_runs}，单倍段落={spacing_paragraphs}/{paragraphs}，"
        f"不拆行={cant_split_rows}/{total_rows}，重复表头={header_tables}/{len(tables)}"
    )


def margin_check(document: Document) -> tuple[bool, str]:
    values = []
    passed = True
    for section in document.sections:
        current = tuple(round(value.cm, 4) for value in (section.top_margin, section.bottom_margin, section.left_margin, section.right_margin))
        values.append(current)
        targets = (2.54, 2.54, 3.18, 3.18)
        if any(abs(value - target) > 0.05 for value, target in zip(current, targets)):
            passed = False
    return passed and bool(values), f"节数={len(values)}，页边距={values}"


def render_check(path: Path) -> tuple[bool, str]:
    libreoffice = shutil.which("libreoffice")
    pdfinfo = shutil.which("pdfinfo")
    pdftotext = shutil.which("pdftotext")
    if not all((libreoffice, pdfinfo, pdftotext)):
        return False, "缺少 libreoffice/pdfinfo/pdftotext"
    with tempfile.TemporaryDirectory(prefix="cowork-evo-018-") as temporary:
        temp_path = Path(temporary)
        profile = temp_path / "lo-profile"
        environment = os.environ.copy()
        for key in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
            environment.pop(key, None)
        completed = subprocess.run(
            [
                libreoffice,
                f"-env:UserInstallation={profile.as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(temp_path),
                str(path),
            ],
            capture_output=True,
            text=True,
            timeout=600,
            check=False,
            env=environment,
        )
        pdf_path = temp_path / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf_path.is_file():
            return False, (completed.stderr or completed.stdout or "转换失败")[-1000:]
        info = subprocess.run([pdfinfo, str(pdf_path)], capture_output=True, text=True, timeout=60, check=False)
        text_result = subprocess.run([pdftotext, "-layout", str(pdf_path), "-"], capture_output=True, text=True, timeout=180, check=False)
        page_match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        pages = int(page_match.group(1)) if page_match else 0
        normalized = normalize_text(text_result.stdout)
        required = [TITLE, "摘要", "七、结语与行动清单", "图31", "表25"]
        missing = [value for value in required if normalize_text(value) not in normalized]
        low, high = EXPECTATIONS["render_page_range"]
        passed = info.returncode == 0 and text_result.returncode == 0 and low <= pages <= high and not missing
        return passed, f"页数={pages}，缺失={missing}，pdf={pdf_path.stat().st_size if pdf_path.exists() else 0} bytes"


def duplicate_heading_prefixes(document: Document) -> int:
    patterns = [
        re.compile(r"^[一二三四五六七八九十]+、[一二三四五六七八九十]+、"),
        re.compile(r"^（[一二三四五六七八九十]+）\s*（[一二三四五六七八九十]+）"),
        re.compile(r"^\d+\.\s*\d+\."),
        re.compile(r"^（\d+）\s*（\d+）"),
    ]
    return sum(any(pattern.search(paragraph.text.strip()) for pattern in patterns) for paragraph in document.paragraphs)


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)
    try:
        candidate = Document(output_path)
    except Exception as exc:
        results.append(criterion("R001", False, f"文件无法打开：{type(exc).__name__}: {exc}"))
        return build_result(results, output_path)
    results.append(criterion("R001", True, f"DOCX 可打开，{len(candidate.paragraphs)} 段，{len(candidate.tables)} 表"))

    source_path = BASE_DIR / "data" / "input_files" / INPUT_NAME
    source_hash_ok = source_path.is_file() and sha256_file(source_path) == EXPECTATIONS["source_sha256"]
    if not source_hash_ok:
        fidelity_ok, fidelity = False, {"error": "基准输入缺失或 SHA256 不一致"}
    else:
        source = Document(source_path)
        fidelity_ok, fidelity = document_content_fidelity(source, candidate)
    results.append(criterion("R002", fidelity_ok, json.dumps(fidelity, ensure_ascii=False, sort_keys=True)))

    title_ok, title_evidence = title_and_abstract_check(candidate)
    results.append(criterion("R003", title_ok, title_evidence))
    hierarchy_ok, hierarchy_evidence, grouped = heading_check(candidate)
    results.append(criterion("R004", hierarchy_ok, hierarchy_evidence))
    heading_fonts_ok, heading_fonts_evidence = heading_font_check(candidate, grouped)
    results.append(criterion("R005", heading_fonts_ok, heading_fonts_evidence))
    body_ok, body_evidence = body_format_check(candidate)
    results.append(criterion("R006", body_ok, body_evidence))

    figure_ok, figure_evidence = caption_check(candidate, "figure")
    table_caption_ok, table_caption_evidence = caption_check(candidate, "table")
    results.append(criterion("R007", figure_ok, json.dumps(figure_evidence, ensure_ascii=False, sort_keys=True)))
    results.append(criterion("R008", table_caption_ok, json.dumps(table_caption_evidence, ensure_ascii=False, sort_keys=True)))

    tables_ok, tables_evidence = table_format_check(candidate)
    results.append(criterion("R009", tables_ok, tables_evidence))
    margins_ok, margins_evidence = margin_check(candidate)
    results.append(criterion("R010", margins_ok, margins_evidence))
    rendered_ok, rendered_evidence = render_check(output_path)
    results.append(criterion("R011", rendered_ok, rendered_evidence))

    missing_captions = max(0, EXPECTATIONS["object_counts"]["figure_captions"] - figure_evidence["captions"]) + max(
        0, EXPECTATIONS["object_counts"]["table_captions"] - table_caption_evidence["captions"]
    )
    results.append(criterion("R012", missing_captions >= 5, f"缺失题注={missing_captions}"))
    severe_loss = (
        fidelity.get("body_block_ratio", 0) < 0.80
        or fidelity.get("equal_tables", 0) < 20
        or fidelity.get("charts", 0) < 15
    )
    results.append(criterion("R013", severe_loss, json.dumps(fidelity, ensure_ascii=False, sort_keys=True)))
    placement_errors = (
        figure_evidence["targets"] - figure_evidence["positioned"]
        + table_caption_evidence["targets"] - table_caption_evidence["positioned"]
    )
    duplicate_prefixes = duplicate_heading_prefixes(candidate)
    results.append(criterion("R014", placement_errors >= 3 or duplicate_prefixes >= 2, f"错位={placement_errors}，重复编号={duplicate_prefixes}"))
    return build_result(results, output_path)


def build_result(results: list[dict], artifact_path: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier, (_, _, kind) in CRITERION_SPECS.items():
        if identifier not in present:
            results.append(criterion(identifier, False, "因文件缺失、损坏或门槛失败未执行"))
    results.sort(key=lambda item: item["criterion_id"])

    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = 0.0

    artifact_digest = sha256_file(artifact_path) if artifact_path.is_file() else hashlib.sha256(b"").hexdigest()
    normalized = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        normalized.append({
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
        "artifact_sha256": [artifact_digest],
        "criterion_results": normalized,
        "reward": round(reward, 6),
        "passed": reward >= PASS_THRESHOLD,
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
