#!/usr/bin/env python3
"""从 Cowork-Bench 固定 SQL dump 中原样抽取本题唯一依赖的 EMPLOYEES COPY 块。

脚本不生成、重排或补造任何员工记录；输出是解压后 SQL 中从
COPY 声明到 ``\.`` 终止行的连续原始字节切片。聚合校验值来自同一
固定 dump 与上游 Evaluator，用来防止抽取错表或截断。
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
from collections import defaultdict
from decimal import Decimal
from pathlib import Path


DUMP_SHA256 = "bd9c45b5ec9072b9f301b7df5b3e9ac65540ffcd53485329ebea44b0c3710da4"
SNAPSHOT_SHA256 = "c1dfba6c1b0c55b03e47cb46d093202af2befb9593bc7f4cf4796a9781844993"
COPY_PREFIX = b'COPY sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES" '
EXPECTED_ROWS = 50_000
EXPECTED_AGGREGATES = {
    "Инженерия": (7096, Decimal("418604451.00"), Decimal("58991.61")),
    "Кадры": (7077, Decimal("416980025.00"), Decimal("58920.45")),
    "НИОКР": (7083, Decimal("410147671.00"), Decimal("57905.93")),
    "Операции": (7120, Decimal("411598247.00"), Decimal("57808.74")),
    "Поддержка": (7244, Decimal("423053067.00"), Decimal("58400.48")),
    "Продажи": (7232, Decimal("425710161.00"), Decimal("58864.79")),
    "Финансы": (7148, Decimal("413713333.00"), Decimal("57878.19")),
}


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def extract(dump_path: Path, output_path: Path) -> None:
    if digest(dump_path) != DUMP_SHA256:
        raise ValueError("上游 db/init.sql.gz 不是记录的固定版本")

    payload: list[bytes] = []
    active = False
    with gzip.open(dump_path, "rb") as stream:
        for line in stream:
            if not active:
                if line.startswith(COPY_PREFIX):
                    active = True
                    payload.append(line)
                continue
            payload.append(line)
            if line.rstrip(b"\r\n") == b"\\.":
                break

    if not payload or payload[-1].rstrip(b"\r\n") != b"\\.":
        raise ValueError("未找到完整的 EMPLOYEES COPY 块")

    data_lines = payload[1:-1]
    if len(data_lines) != EXPECTED_ROWS:
        raise ValueError(f"行数异常：{len(data_lines)}")

    aggregates: dict[str, list[Decimal | int]] = defaultdict(lambda: [0, Decimal("0")])
    for line_number, raw in enumerate(data_lines, start=1):
        fields = raw.rstrip(b"\r\n").decode("utf-8").split("\t")
        if len(fields) != 18:
            raise ValueError(f"第 {line_number} 行字段数不是 18：{len(fields)}")
        department = fields[11]
        salary = Decimal(fields[4])
        aggregates[department][0] += 1
        aggregates[department][1] += salary

    actual = {
        department: (
            int(values[0]),
            Decimal(values[1]).quantize(Decimal("0.01")),
            (Decimal(values[1]) / int(values[0])).quantize(Decimal("0.01")),
        )
        for department, values in aggregates.items()
    }
    if actual != EXPECTED_AGGREGATES:
        raise ValueError(f"部门聚合不符合固定上游快照：{actual}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(b"".join(payload))
    snapshot_hash = digest(output_path)
    if snapshot_hash != SNAPSHOT_SHA256:
        raise ValueError(f"输出快照哈希异常：{snapshot_hash}")
    print(f"{output_path}: {len(data_lines)} 行，sha256={snapshot_hash}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dump", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    extract(args.dump, args.output)


if __name__ == "__main__":
    main()
