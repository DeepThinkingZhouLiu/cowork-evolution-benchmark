#!/usr/bin/env python3
"""Cowork-Bench RZD 团队出差计划题的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

from docx import Document


OUTPUT_NAME = "Travel_Plan.docx"
TASK_ID = "cowork-evo/cowork-bench-rzd-trip-plan-from-outline"
RUBRIC_VERSION = "cowork-bench-rzd-trip-plan-static-v0.1"
JUDGE_VERSION = "cowork-bench-rzd-trip-plan-static-v0.1"
HURDLE_CAPS = {"R001": 0.0, **{f"R{index:03d}": 0.74 for index in range(2, 7)}}

CRITERION_SPECS = {
    "R001": ("指定 DOCX 可打开、可编辑且包含规定章节与表格", 4, "hurdle"),
    "R002": ("去程 752А 的车站、时刻和时长全部正确", 6, "hurdle"),
    "R003": ("返程 759А 的车站、时刻和时长全部正确", 6, "hurdle"),
    "R004": ("按三人团队和预算正确选择经济舱", 5, "hurdle"),
    "R005": ("单票、单程团队价和往返总额全部正确", 6, "hurdle"),
    "R006": ("预算和余额计算正确", 4, "hurdle"),
    "R007": ("日期、团队规模和舱位摘要完整", 3, "positive"),
    "R008": ("章节顺序、表字段和真实 Word 表格符合指南", 4, "positive"),
    "R009": ("文档可由 LibreOffice 渲染且关键内容可读", 3, "positive"),
    "R010": ("任一车次、时刻或车站错误", -10, "penalty"),
    "R011": ("任一舱位、票价、总额、预算或余额错误", -10, "penalty"),
    "R012": ("文件损坏、截图化或无法渲染", -10, "penalty"),
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


def normalized(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def train_normalized(value) -> str:
    text = unicodedata.normalize("NFKD", normalized(value))
    return text.translate(str.maketrans({"а": "a", "о": "o", "е": "e", "р": "p", "с": "c", "у": "y", "к": "k", "х": "x"}))


def money(value) -> int | None:
    match = re.search(r"(\d[\d\s,.]*)", str(value or ""))
    if not match:
        return None
    digits = re.sub(r"\D", "", match.group(1))
    return int(digits) if digits else None


def minutes(value: str) -> int:
    hour, minute = value.split(":")
    return int(hour) * 60 + int(minute)


def derive_expected(input_dir: Path):
    config = json.loads((input_dir / "trip_config.json").read_text(encoding="utf-8"))
    snapshot = json.loads((input_dir / "travel_snapshot.json").read_text(encoding="utf-8"))
    meeting_start = datetime.fromisoformat(snapshot["client_meeting"]["start"])
    meeting_end = datetime.fromisoformat(snapshot["client_meeting"]["end"])
    latest_arrival = meeting_start.hour * 60 + meeting_start.minute - 60
    outbound = min(
        (train for train in snapshot["outbound_trains"] if minutes(train["arrival"]) <= latest_arrival),
        key=lambda train: minutes(train["arrival"]),
    )
    earliest_return = meeting_end.hour * 60 + meeting_end.minute + 30
    returning = min(
        (train for train in snapshot["return_trains"] if minutes(train["departure"]) >= earliest_return),
        key=lambda train: minutes(train["departure"]),
    )
    team = int(config["team_size"])
    budget = int(config["budget_total_rub"])
    classes = sorted(outbound["fares"], key=lambda item: outbound["fares"][item] + returning["fares"][item])
    seat_class = next(
        item for item in classes
        if (outbound["fares"][item] + returning["fares"][item]) * team <= budget
    )
    outbound_team = int(outbound["fares"][seat_class] * team)
    return_team = int(returning["fares"][seat_class] * team)
    total = outbound_team + return_team
    return {
        "config": config,
        "snapshot": snapshot,
        "outbound": outbound,
        "returning": returning,
        "seat_class": seat_class,
        "ticket": int(outbound["fares"][seat_class]),
        "outbound_team": outbound_team,
        "return_team": return_team,
        "total": total,
        "budget": budget,
        "remainder": budget - total,
    }


def table_mapping(table) -> tuple[dict[str, str], list[str]]:
    mapping = {}
    labels = []
    for row in table.rows:
        if len(row.cells) < 2:
            continue
        label = normalized(row.cells[0].text)
        labels.append(label)
        mapping[label] = row.cells[1].text.strip()
    return mapping, labels


def lookup(mapping: dict[str, str], fragment: str):
    fragment = normalized(fragment)
    matches = [value for key, value in mapping.items() if fragment in key]
    return matches[0] if len(matches) == 1 else None


def libreoffice_render(path: Path, root: Path):
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    if libreoffice is None or pdftotext is None:
        return "", "缺少 libreoffice 或 pdftotext"
    source_dir = root / "source"
    pdf_dir = root / "pdf"
    source_dir.mkdir()
    pdf_dir.mkdir()
    source_copy = source_dir / OUTPUT_NAME
    shutil.copy2(path, source_copy)
    profile_uri = (root / "lo-profile").resolve().as_uri()
    process = subprocess.run(
        [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(source_copy)],
        capture_output=True, text=True, timeout=300, check=False,
    )
    pdf_path = pdf_dir / "Travel_Plan.pdf"
    if process.returncode != 0 or not pdf_path.is_file():
        return "", (process.stderr or process.stdout or "LibreOffice PDF 渲染失败")[-1000:]
    extracted = subprocess.run(
        [pdftotext, "-raw", str(pdf_path), "-"],
        capture_output=True, text=True, timeout=120, check=False,
    )
    if extracted.returncode != 0:
        return "", (extracted.stderr or "PDF 文本提取失败")[-1000:]
    return extracted.stdout, "LibreOffice PDF 渲染和文本提取成功"


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)

    expected = derive_expected(Path(__file__).resolve().parents[1] / "data" / "input_files")
    with tempfile.TemporaryDirectory(prefix="cowork-evo-rzd-word-") as temp_dir:
        rendered_text, render_evidence = libreoffice_render(output_path, Path(temp_dir))
        try:
            if not zipfile.is_zipfile(output_path):
                raise ValueError("不是有效 OOXML ZIP")
            document = Document(output_path)
            paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
            headings = [
                paragraph.text.strip()
                for paragraph in document.paragraphs
                if paragraph.text.strip() and paragraph.style.name.startswith("Heading")
            ]
            tables = [table_mapping(table) for table in document.tables]
            full_text = "\n".join(paragraphs + [cell.text for table in document.tables for row in table.rows for cell in row.cells])
            with zipfile.ZipFile(output_path) as archive:
                media_count = sum(1 for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/"))
        except Exception as error:
            results.append(criterion("R001", False, f"DOCX 无法读取：{error}"))
            return build_result(results, output_path)

    required_headings = ["Поездка туда", "Поездка обратно", "Сводка по билетам и бюджет"]
    structure_ok = (
        len(tables) == 3
        and [normalized(item) for item in headings] == [normalized(item) for item in required_headings]
        and paragraphs
        and normalized(paragraphs[0]) == normalized("План деловой поездки команды Москва — Санкт-Петербург")
        and media_count == 0
    )
    results.append(criterion("R001", structure_ok, f"标题={paragraphs[:1]}，章节={headings}，表格={len(tables)}，媒体={media_count}"))

    outbound_map = tables[0][0] if len(tables) > 0 else {}
    return_map = tables[1][0] if len(tables) > 1 else {}
    summary_map = tables[2][0] if len(tables) > 2 else {}
    stations = expected["snapshot"]["stations"]

    outbound_train = lookup(outbound_map, "номер поезда")
    outbound_departure = lookup(outbound_map, "время отправления")
    outbound_arrival = lookup(outbound_map, "время прибытия")
    outbound_duration = lookup(outbound_map, "длительность")
    outbound_ok = (
        "752a" in train_normalized(outbound_train)
        and normalized(outbound_departure) == expected["outbound"]["departure"]
        and normalized(outbound_arrival) == expected["outbound"]["arrival"]
        and normalized(outbound_duration) == expected["outbound"]["duration"]
        and normalized(stations[expected["outbound"]["from"]]) in normalized(full_text)
        and normalized(stations[expected["outbound"]["to"]]) in normalized(full_text)
    )
    results.append(
        criterion(
            "R002", outbound_ok,
            f"车次={outbound_train}，出发={outbound_departure}，到达={outbound_arrival}，时长={outbound_duration}，车站存在={normalized(stations['MOW']) in normalized(full_text)}/{normalized(stations['SPB']) in normalized(full_text)}",
        )
    )

    return_train = lookup(return_map, "номер поезда")
    return_departure = lookup(return_map, "время отправления")
    return_arrival = lookup(return_map, "время прибытия")
    return_duration = lookup(return_map, "длительность")
    return_ok = (
        "759a" in train_normalized(return_train)
        and normalized(return_departure) == expected["returning"]["departure"]
        and normalized(return_arrival) == expected["returning"]["arrival"]
        and normalized(return_duration) == expected["returning"]["duration"]
        and normalized(stations[expected["returning"]["from"]]) in normalized(full_text)
        and normalized(stations[expected["returning"]["to"]]) in normalized(full_text)
    )
    results.append(criterion("R003", return_ok, f"车次={return_train}，出发={return_departure}，到达={return_arrival}，时长={return_duration}"))

    class_ok = (
        "эконом" in normalized(full_text)
        and "эконом+" not in normalized(full_text)
        and "бизнес" not in normalized(full_text)
    )
    results.append(criterion("R004", class_ok, f"经济舱出现={'эконом' in normalized(full_text)}，经济+出现={'эконом+' in normalized(full_text)}，商务舱出现={'бизнес' in normalized(full_text)}"))

    outbound_ticket = lookup(outbound_map, "цена за один билет")
    return_ticket = lookup(return_map, "цена за один билет")
    outbound_team = lookup(outbound_map, "стоимость на команду")
    return_team = lookup(return_map, "стоимость на команду")
    total = lookup(summary_map, "итого на команду")
    prices_ok = (
        money(outbound_ticket) == expected["ticket"]
        and money(return_ticket) == expected["ticket"]
        and money(outbound_team) == expected["outbound_team"]
        and money(return_team) == expected["return_team"]
        and money(total) == expected["total"]
    )
    results.append(
        criterion(
            "R005", prices_ok,
            f"单票={money(outbound_ticket)}/{money(return_ticket)}，团队单程={money(outbound_team)}/{money(return_team)}，总额={money(total)}",
        )
    )

    budget = lookup(summary_map, "бюджет")
    remainder = lookup(summary_map, "остаток после поездки")
    budget_ok = money(budget) == expected["budget"] and money(remainder) == expected["remainder"]
    results.append(criterion("R006", budget_ok, f"预算={money(budget)}，余额={money(remainder)}"))

    summary_ok = (
        ("10 марта 2026" in normalized(full_text) or "2026-03-10" in normalized(full_text) or "10.03.2026" in normalized(full_text))
        and re.search(r"(?:команда|команду).{0,20}3|3.{0,20}(?:человек|команда)", normalized(full_text)) is not None
        and "эконом" in normalized(full_text)
    )
    results.append(criterion("R007", summary_ok, f"日期、三人团队、经济舱摘要通过={summary_ok}"))

    expected_outbound_labels = ["номер поезда", "время отправления", "время прибытия", "длительность в пути", "цена за один билет (эконом)", "стоимость на команду из 3 человек"]
    expected_return_labels = expected_outbound_labels
    expected_summary_labels = ["туда на команду", "обратно на команду", "итого на команду", "бюджет", "остаток после поездки"]
    labels_ok = (
        len(tables) == 3
        and tables[0][1] == expected_outbound_labels
        and tables[1][1] == expected_return_labels
        and tables[2][1] == expected_summary_labels
        and [normalized(item) for item in headings] == [normalized(item) for item in required_headings]
    )
    results.append(criterion("R008", labels_ok, f"去程标签={tables[0][1] if len(tables)>0 else []}，返程标签={tables[1][1] if len(tables)>1 else []}，汇总标签={tables[2][1] if len(tables)>2 else []}"))

    rendered = train_normalized(rendered_text)
    render_ok = all(token in rendered for token in ("752a", "759a", "33 000", "40 000"))
    results.append(criterion("R009", render_ok, f"{render_evidence}；渲染字符={len(rendered)}"))

    itinerary_error = not (outbound_ok and return_ok)
    budget_error = not (class_ok and prices_ok and budget_ok)
    broken = not (structure_ok and render_ok) or media_count > 0
    results.append(criterion("R010", itinerary_error, f"去程={outbound_ok}，返程={return_ok}"))
    results.append(criterion("R011", budget_error, f"舱位={class_ok}，金额={prices_ok}，预算余额={budget_ok}"))
    results.append(criterion("R012", broken, f"结构={structure_ok}，渲染={render_ok}，媒体={media_count}"))
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
    hurdle_failures = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
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
