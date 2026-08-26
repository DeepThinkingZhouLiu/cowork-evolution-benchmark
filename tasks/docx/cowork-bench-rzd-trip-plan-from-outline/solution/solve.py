#!/usr/bin/env python3
"""Cowork-Bench RZD 团队出差计划题的逐题参考解法。"""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


def minutes(value: str) -> int:
    hour, minute = value.split(":")
    return int(hour) * 60 + int(minute)


def select_plan(input_dir: Path):
    config = json.loads((input_dir / "trip_config.json").read_text(encoding="utf-8"))
    snapshot = json.loads((input_dir / "travel_snapshot.json").read_text(encoding="utf-8"))
    meeting_start = datetime.fromisoformat(snapshot["client_meeting"]["start"])
    meeting_end = datetime.fromisoformat(snapshot["client_meeting"]["end"])

    # 30 分钟市内交通 + 至少 30 分钟准备；13:25 到达的 754А 只有 35 分钟，不满足。
    latest_reasonable_arrival = meeting_start.hour * 60 + meeting_start.minute - 60
    outbound_candidates = [
        train for train in snapshot["outbound_trains"]
        if minutes(train["arrival"]) <= latest_reasonable_arrival
    ]
    outbound = min(outbound_candidates, key=lambda train: minutes(train["arrival"]))

    earliest_reasonable_return = meeting_end.hour * 60 + meeting_end.minute + 30
    return_candidates = [
        train for train in snapshot["return_trains"]
        if minutes(train["departure"]) >= earliest_reasonable_return
    ]
    returning = min(return_candidates, key=lambda train: minutes(train["departure"]))

    team_size = int(config["team_size"])
    budget = int(config["budget_total_rub"])
    classes = sorted(
        outbound["fares"],
        key=lambda seat_class: outbound["fares"][seat_class] + returning["fares"][seat_class],
    )
    seat_class = next(
        seat_class for seat_class in classes
        if (outbound["fares"][seat_class] + returning["fares"][seat_class]) * team_size <= budget
    )
    outbound_team = int(outbound["fares"][seat_class] * team_size)
    return_team = int(returning["fares"][seat_class] * team_size)
    total = outbound_team + return_team
    return config, snapshot, outbound, returning, seat_class, outbound_team, return_team, total, budget - total


def rubles(value: int) -> str:
    return f"{value:,}".replace(",", " ") + " ₽"


def add_two_column_table(document: Document, rows: list[tuple[str, str]]) -> None:
    table = document.add_table(rows=0, cols=2)
    table.style = "Light Grid Accent 1"
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
        for run in cells[0].paragraphs[0].runs:
            run.bold = True


def solve(input_dir: Path, output_dir: Path) -> None:
    config, snapshot, outbound, returning, seat_class, outbound_team, return_team, total, remainder = select_plan(input_dir)
    stations = snapshot["stations"]

    document = Document()
    styles = document.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10.5)
    title = document.add_heading("План деловой поездки команды Москва — Санкт-Петербург", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_paragraph(
        "Дата поездки: 10 марта 2026 года. "
        f"Команда: {config['team_size']} человека. Класс обслуживания: {seat_class}."
    )

    document.add_heading("Поездка туда", level=1)
    document.add_paragraph(
        f"Выбранный поезд из Москвы ({stations[outbound['from']]}) "
        f"в Санкт-Петербург ({stations[outbound['to']]}):"
    )
    add_two_column_table(
        document,
        [
            ("Номер поезда", f"{outbound['train_code']} (Сапсан)"),
            ("Время отправления", outbound["departure"]),
            ("Время прибытия", outbound["arrival"]),
            ("Длительность в пути", outbound["duration"]),
            (f"Цена за один билет ({seat_class})", rubles(int(outbound["fares"][seat_class]))),
            (f"Стоимость на команду из {config['team_size']} человек", rubles(outbound_team)),
        ],
    )

    document.add_heading("Поездка обратно", level=1)
    document.add_paragraph(
        f"Выбранный поезд из Санкт-Петербурга ({stations[returning['from']]}) "
        f"в Москву ({stations[returning['to']]}):"
    )
    add_two_column_table(
        document,
        [
            ("Номер поезда", f"{returning['train_code']} (Сапсан)"),
            ("Время отправления", returning["departure"]),
            ("Время прибытия", returning["arrival"]),
            ("Длительность в пути", returning["duration"]),
            (f"Цена за один билет ({seat_class})", rubles(int(returning["fares"][seat_class]))),
            (f"Стоимость на команду из {config['team_size']} человек", rubles(return_team)),
        ],
    )

    document.add_heading("Сводка по билетам и бюджет", level=1)
    add_two_column_table(
        document,
        [
            ("Туда на команду", rubles(outbound_team)),
            ("Обратно на команду", rubles(return_team)),
            ("Итого на команду", rubles(total)),
            ("Бюджет", rubles(int(config["budget_total_rub"]))),
            ("Остаток после поездки", rubles(remainder)),
        ],
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    document.save(output_dir / "Travel_Plan.docx")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
