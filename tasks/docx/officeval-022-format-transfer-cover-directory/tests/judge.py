#!/usr/bin/env python3
"""officeval_022 的 Harbor Judge。

保留上游 60 分版式规则，同时补充源事实保真、目录实际页码、字段页脚、正文可编辑性和
LibreOffice 真渲染检查，防止只针对样式检测造壳文档。
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
from docx.oxml.ns import qn

import upstream_verifier


OUTPUT_ONE = "企业经营情况说明书_按模板格式添加封面目录.docx"
OUTPUT_TWO = "华中区域零售项目经营资金异常及现金流风险_主题修改版.docx"
TASK_ID = "cowork-evo/officeval-022-format-transfer-cover-directory"
RUBRIC_VERSION = "omegause-officeval-022-adaptation-v0.1"
JUDGE_VERSION = "officeval-022-adaptation-v0.1"

CRITERION_SPECS = {
    "R001": ("两份指定名称的 Word 均可打开、可编辑，且不是页面截图替代物", 3, "hurdle"),
    "R002": ("第一份 Word 完整保留华中区域说明书的真实正文、四张表和资金流转图", 5, "positive"),
    "R003": ("新增封面逐项使用给定封面事实，三条 7.7–8.1 厘米虚线及版式正确", 4, "positive"),
    "R004": ("目录含六个真实一级标题，点引导与页码准确对应实际渲染页", 4, "positive"),
    "R005": ("第一份正文标题、正文、图注和三线表满足上游格式契约", 5, "positive"),
    "R006": ("第二份只替换指定主题并完整保留原报告事实、表格和签名", 5, "positive"),
    "R007": ("第二份标题层级、正文公司名、三线表与签名格式没有触发上游扣分", 4, "positive"),
    "R008": ("两份 Word 均能真实渲染为可读 A4，多页内容无空页，第一页目录页脚字段可重算", 3, "positive"),
    "R009": ("任一报告丢失两个及以上关键事实、表格或图片", -8, "penalty"),
    "R010": ("任一文档以整页图片或不可编辑对象替代主体文字", -10, "penalty"),
    "R011": ("文件角色颠倒、第二份主题错误，或第一份六个正文标题不完整", -8, "penalty"),
    "R012": ("目录页码与实际分页不一致、页脚字段缺失或任一文档无法渲染", -6, "penalty"),
}


def normalized(text: str) -> str:
    return re.sub(r"\s+", "", text or "").replace("：", ":").replace("（", "(").replace("）", ")")


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


def document_text(document: Document) -> str:
    parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return normalized("\n".join(parts))


def editable_stats(path: Path) -> tuple[Document, dict]:
    document = Document(path)
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml").decode("utf-8", errors="replace")
        media = [name for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/")]
    regular_chars = len(normalized("\n".join(p.text for p in document.paragraphs)))
    drawing_count = xml.count("<w:drawing") + xml.count("<w:pict")
    return document, {
        "regular_chars": regular_chars,
        "drawing_count": drawing_count,
        "media_count": len(media),
        "tables": len(document.tables),
    }


def toc_state(document: Document, expected: dict) -> dict:
    headings = expected["first_headings"]
    toc_entries = {}
    for paragraph in document.paragraphs:
        compact = normalized(paragraph.text)
        for heading in headings:
            h = normalized(heading)
            if compact.startswith(h) and re.search(r"\d+$", compact):
                toc_entries[heading] = int(re.search(r"(\d+)$", compact).group(1))
    footer_xml = "".join(section.footer._element.xml for section in document.sections)
    return {
        "entries": toc_entries,
        "has_page_field": "PAGE" in footer_xml,
        "has_numpages_field": "NUMPAGES" in footer_xml,
    }


def render_documents(paths: list[Path], headings: list[str]) -> dict:
    libreoffice = shutil.which("libreoffice")
    pdfinfo = shutil.which("pdfinfo")
    pdftotext = shutil.which("pdftotext")
    if not libreoffice or not pdfinfo or not pdftotext:
        return {"ok": False, "error": "缺少 libreoffice、pdfinfo 或 pdftotext"}
    details = {}
    try:
        with tempfile.TemporaryDirectory(prefix="cowork-evo-022-") as temp_dir:
            root = Path(temp_dir)
            profile_uri = (root / "lo-profile").resolve().as_uri()
            for path in paths:
                converted = subprocess.run(
                    [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}",
                     "--convert-to", "pdf", "--outdir", temp_dir, str(path)],
                    capture_output=True, text=True, timeout=300, check=False,
                )
                pdf = root / f"{path.stem}.pdf"
                if converted.returncode != 0 or not pdf.is_file():
                    return {"ok": False, "error": (converted.stderr or converted.stdout)[-1000:]}
                info = subprocess.run([pdfinfo, str(pdf)], capture_output=True, text=True,
                                      timeout=120, check=False).stdout
                pages_match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
                size_match = re.search(r"^Page size:\s+([\d.]+) x ([\d.]+)", info, re.MULTILINE)
                pages = int(pages_match.group(1)) if pages_match else 0
                page_texts = []
                for page_number in range(1, pages + 1):
                    page_text = subprocess.run(
                        [pdftotext, "-f", str(page_number), "-l", str(page_number),
                         "-layout", str(pdf), "-"],
                        capture_output=True, text=True, timeout=120, check=False,
                    ).stdout
                    page_texts.append(page_text)
                width = float(size_match.group(1)) if size_match else 0
                height = float(size_match.group(2)) if size_match else 0
                details[path.name] = {
                    "pages": pages,
                    "page_chars": [len(normalized(text)) for text in page_texts],
                    "a4": abs(width - 595.3) < 8 and abs(height - 841.9) < 8,
                    "heading_pages": {
                        heading: next(
                            (index + 1 for index, text in enumerate(page_texts)
                             if index != 1 and normalized(heading) in normalized(text)),
                            None,
                        )
                        for heading in headings
                    } if path.name == OUTPUT_ONE else {},
                }
    except Exception as exc:
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    readable = all(
        value["pages"] >= 2 and value["a4"] and min(value["page_chars"]) >= 25
        for value in details.values()
    )
    return {"ok": readable, "details": details}


def upstream_state(artifact_dir: Path) -> dict:
    result = upstream_verifier.evaluate(str(artifact_dir))
    misses = [item["rule"] for item in result.get("dim2_items", [])
              if item.get("max_delta", 0) > 0 and not item.get("hit")]
    penalties = [item["rule"] for item in result.get("dim2_items", [])
                 if item.get("delta", 0) < 0 and item.get("hit")]
    return {
        "dim1": result.get("dim1_pass", False),
        "score": result.get("total_score", 0),
        "max": result.get("max_score", 0),
        "misses": misses,
        "penalties": penalties,
    }


def evaluate(artifact_dir: Path) -> dict:
    expected = json.loads(Path(__file__).with_name("expectations.json").read_text(encoding="utf-8"))
    first_path = artifact_dir / OUTPUT_ONE
    second_path = artifact_dir / OUTPUT_TWO
    results = []
    if not first_path.is_file() or not second_path.is_file():
        missing = [name for name, path in [(OUTPUT_ONE, first_path), (OUTPUT_TWO, second_path)] if not path.is_file()]
        results.append(criterion("R001", False, f"缺少文件={missing}"))
        return build_result(results, [first_path, second_path])
    try:
        first, first_stats = editable_stats(first_path)
        second, second_stats = editable_stats(second_path)
    except Exception as exc:
        results.append(criterion("R001", False, f"Word 无法打开：{type(exc).__name__}: {exc}"))
        return build_result(results, [first_path, second_path])

    editable = (
        first_stats["regular_chars"] >= 1800 and second_stats["regular_chars"] >= 2400
        and first_stats["drawing_count"] <= 5 and second_stats["drawing_count"] <= 2
    )
    results.append(criterion("R001", editable, f"第一份={first_stats}，第二份={second_stats}"))

    first_text = document_text(first)
    second_text = document_text(second)
    first_missing = [item for item in expected["first_source_facts"] if normalized(item) not in first_text]
    first_structure = first_stats["tables"] == 4 and first_stats["media_count"] >= 1
    results.append(criterion("R002", not first_missing and first_structure,
                             f"缺失事实={first_missing}，表格={first_stats['tables']}，媒体={first_stats['media_count']}"))

    cover_missing = [item for item in expected["cover_fragments"] if normalized(item) not in first_text]
    upstream = upstream_state(artifact_dir)
    cover_rules_missed = [item for item in upstream["misses"] if "封面" in item or "分隔" in item]
    results.append(criterion("R003", not cover_missing and not cover_rules_missed,
                             f"缺失封面事实={cover_missing}，上游封面未命中={cover_rules_missed}"))

    toc = toc_state(first, expected)
    render = render_documents([first_path, second_path], expected["first_headings"])
    rendered_first = render.get("details", {}).get(OUTPUT_ONE, {})
    actual_heading_pages = rendered_first.get("heading_pages", {})
    toc_consistent = (
        len(toc["entries"]) == 6 and actual_heading_pages
        and all(toc["entries"].get(h) == actual_heading_pages.get(h) for h in expected["first_headings"])
    )
    results.append(criterion("R004", bool(toc_consistent),
                             f"目录={toc['entries']}，实际页={actual_heading_pages}"))

    first_headings = [p.text.strip() for p in first.paragraphs if p.style and p.style.name == "Heading 1"]
    first_subheadings = [p.text.strip() for p in first.paragraphs if p.style and p.style.name == "Heading 2"]
    first_format_ok = upstream["score"] == upstream["max"] == 60
    results.append(criterion("R005", first_format_ok and first_headings == expected["first_headings"]
                             and first_subheadings == expected["first_subheadings"],
                             f"上游={upstream['score']}/{upstream['max']}，一级={first_headings}，二级={first_subheadings}"))

    second_missing = [item for item in expected["second_source_facts"] if normalized(item) not in second_text]
    title_text = second.paragraphs[0].text.strip() if second.paragraphs else ""
    second_structure = second_stats["tables"] == 3
    results.append(criterion("R006", title_text == expected["second_title"] and not second_missing and second_structure,
                             f"标题={title_text!r}，缺失事实={second_missing}，表格={second_stats['tables']}"))
    results.append(criterion("R007", not upstream["penalties"],
                             f"上游扣分项={upstream['penalties']}"))

    footer_ok = toc["has_page_field"] and toc["has_numpages_field"]
    results.append(criterion("R008", render.get("ok", False) and footer_ok,
                             f"渲染={render}，PAGE={toc['has_page_field']}，NUMPAGES={toc['has_numpages_field']}"))

    content_error_count = len(first_missing) + len(second_missing)
    structural_loss = first_stats["tables"] < 4 or second_stats["tables"] < 3 or first_stats["media_count"] < 1
    results.append(criterion("R009", content_error_count >= 2 or structural_loss,
                             f"关键事实缺失={content_error_count}，结构丢失={structural_loss}"))
    flattened = first_stats["regular_chars"] < 900 or second_stats["regular_chars"] < 1200 \
        or first_stats["drawing_count"] > 8 or second_stats["drawing_count"] > 5
    results.append(criterion("R010", flattened,
                             f"第一份={first_stats}，第二份={second_stats}"))
    role_error = title_text != expected["second_title"] or first_headings != expected["first_headings"]
    results.append(criterion("R011", role_error,
                             f"第二份标题正确={title_text == expected['second_title']}，第一份标题={first_headings}"))
    render_or_page_error = not render.get("ok", False) or not toc_consistent or not footer_ok
    results.append(criterion("R012", render_or_page_error,
                             f"渲染成功={render.get('ok', False)}，目录一致={toc_consistent}，页脚字段={footer_ok}"))
    return build_result(results, [first_path, second_path])


def build_result(results: list[dict], artifact_paths: list[Path]) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [item["criterion_id"] for item in results
                       if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = 0.0
    digests = [hashlib.sha256(path.read_bytes()).hexdigest() for path in artifact_paths if path.is_file()]
    if not digests:
        digests = [hashlib.sha256(b"").hexdigest()]
    normalized_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        normalized_results.append({
            "criterion_id": item["criterion_id"],
            "score": score,
            "evidence": item["evidence"],
            "method": item["method"],
            "raw": {"description": item["description"], "weight": item["weight"],
                    "type": item["type"], "passed": item["passed"]},
        })
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": digests,
        "criterion_results": normalized_results,
        "reward": round(reward, 6),
        "passed": reward >= 0.75,
        "hurdle_failures": hurdle_failures,
    }


def main():
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
