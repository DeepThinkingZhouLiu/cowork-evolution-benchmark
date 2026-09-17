"""从已保存的 Panel 分数同步难度元数据；不生成题目，也不重新评分。"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import math
import re
import tomllib
from collections import defaultdict
from pathlib import Path


MANIFEST = "benchmarks/difficulty-labels-v0.5.csv"
STATUS = "pilot_single_panel"
SOURCES = {
    "ppt": "benchmarks/ppt-difficulty-pilot-20260904.csv",
    "docx": "benchmarks/docx-difficulty-pilot-20260908.csv",
    "xlsx": "benchmarks/xlsx-difficulty-pilot-20260909.csv",
    "cross_office": "benchmarks/cross-office-hybrid-pilot-20260904.csv",
}
HARNESSES = ("claude-code", "opencode", "openhands")
FIELDS = (
    "instance_id", "task_path", "domain", "rank", "difficulty_level",
    "difficulty_tier", "calibration_status", "claude_code_reward",
    "opencode_reward", "openhands_reward", "panel_mean", "source_report",
    "source_sha256",
)


def rank_label(rank: int) -> tuple[str, str]:
    for end, level, tier in ((12, "L1", "easy"), (30, "L2", "medium"),
                             (48, "L3", "hard"), (60, "L4", "expert")):
        if 1 <= rank <= end:
            return level, tier
    raise ValueError(f"非法名次：{rank}")


def number(value: str) -> float:
    result = float(value)
    if not math.isfinite(result):
        raise ValueError("分数必须为有限数值")
    # Excel 的负向罚分原样保留，不截断为零。
    return result


def source_rows(root: Path, domain: str) -> list[dict]:
    with (root / SOURCES[domain]).open(encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))
    if domain != "cross_office":
        return rows

    grouped: dict[str, dict] = defaultdict(dict)
    for row in rows:
        task, harness = row["task"], row["harness"]
        if harness not in HARNESSES or harness in grouped[task]:
            raise ValueError(f"重复或未知 Harness：{task}/{harness}")
        if row["status"] not in {"completed", "task_failure"}:
            raise ValueError(f"不能将评测基础设施失败用于难度标注：{task}/{harness}")
        score = number(row["combined_reward"])
        code = number(row["deterministic_reward"])
        visual = row["visual_reward"]
        if row["status"] == "task_failure":
            expected = 0.0
        elif row["visual_enabled"] == "true":
            expected = 0.7 * code + 0.3 * number(visual)
        elif row["visual_enabled"] == "false":
            expected = code
        else:
            raise ValueError(f"视觉标记无效：{task}")
        if not math.isclose(score, expected, abs_tol=1e-8):
            raise ValueError(f"混合分与来源公式不一致：{task}/{harness}")
        grouped[task][harness] = score
    ranked = []
    for task, scores in grouped.items():
        if set(scores) != set(HARNESSES):
            raise ValueError(f"缺少三套 Harness 配对结果：{task}")
        ranked.append({"task": task, "claude": scores["claude-code"],
                       "opencode": scores["opencode"], "openhands": scores["openhands"],
                       "panel_mean": sum(scores.values()) / 3})
    ranked.sort(key=lambda row: (-row["panel_mean"], row["task"]))
    if len(ranked) != 60:
        raise ValueError("Cross-Office 必须包含 60 道题")
    for boundary in (12, 30, 48):
        if math.isclose(ranked[boundary - 1]["panel_mean"],
                        ranked[boundary]["panel_mean"], abs_tol=1e-10):
            raise ValueError("Cross-Office 出现边界同分，需逐题结构复核")
    return [dict(row, rank=rank, pilot_tier=rank_label(rank)[1])
            for rank, row in enumerate(ranked, 1)]


def build_manifest(root: Path) -> list[dict]:
    records = []
    for domain, source in SOURCES.items():
        rows = source_rows(root, domain)
        expected = {p.parent.name for p in (root / "tasks" / domain).glob("*/task.toml")}
        if len(rows) != 60 or len(expected) != 60 or {r["task"] for r in rows} != expected:
            raise ValueError(f"{domain} 的来源与当前 60 题未完整一一对应")
        if sorted(int(r["rank"]) for r in rows) != list(range(1, 61)):
            raise ValueError(f"{domain} 名次缺失或重复")
        digest = hashlib.sha256((root / source).read_bytes()).hexdigest()
        previous = math.inf
        for row in sorted(rows, key=lambda r: int(r["rank"])):
            rank = int(row["rank"])
            level, tier = rank_label(rank)
            if row.get("pilot_tier", row.get("candidate_tier", "")).lower() != tier:
                raise ValueError(f"来源档位与名次不符：{domain}/{row['task']}")
            columns = (("claude_code_reward", "opencode_reward", "openhands_reward")
                       if domain == "ppt" else ("claude", "opencode", "openhands"))
            scores = [number(row[column]) for column in columns]
            mean = number(row["panel_mean"])
            if not math.isclose(mean, sum(scores) / 3, abs_tol=1e-8) or mean > previous + 1e-8:
                raise ValueError(f"Panel 均分或排序异常：{domain}/{row['task']}")
            previous = mean
            task_id = f"{domain}/{row['task']}"
            records.append(dict(zip(FIELDS, (
                task_id, f"tasks/{task_id}", domain, rank, level, tier, STATUS,
                *[format(score, ".10f") for score in scores], format(mean, ".10f"),
                source, digest,
            ))))
    return records


def calibration_notes(notes: str) -> str:
    # 只清除过期的待校准状态，保留每份配置自己的来源与制题说明。
    for phrase in (
        "正式难度等待冻结多 Harness Rollout",
        "正式难度等待冻结多Harness Rollout",
        "正式难度待多Harness Rollout",
        "正式难度待 Agent Rollout",
        "difficulty 在冻结多 Harness Rollout 前保持 unassigned",
    ):
        notes = notes.replace(phrase, "当前实测分档见difficulty_source")
    parts = [p.strip() for p in notes.split(";") if p.strip()]
    parts = [p for p in parts if not p.startswith(("calibration_status=", "difficulty_source="))]
    parts = [p for p in parts if p not in {
        "difficulty=unassigned_pending_panel",
        "difficulty remains unassigned",
        "difficulty remains unassigned until frozen panel calibration",
    }]
    return "; ".join(parts + [f"calibration_status={STATUS}", f"difficulty_source={MANIFEST}"])


def metadata_updates(root: Path, record: dict) -> dict[Path, str]:
    """只同步三个元数据字段，保留其他字段及排版。"""
    task = root / record["task_path"]
    path = task / "task.toml"
    text = path.read_text(encoding="utf-8")
    payload = tomllib.loads(text)
    updates = {key: record[key] for key in ("difficulty_level", "difficulty_tier")}
    updates["notes"] = calibration_notes(payload["metadata"].get("notes", ""))
    section = re.search(r"(?ms)^\[metadata\]\s*\n.*?(?=^\[|\Z)", text)
    if section is None:
        raise ValueError(f"缺少 metadata：{path}")
    block = section[0]
    for key, value in updates.items():
        encoded = json.dumps(value, ensure_ascii=False)
        pattern = rf"(?m)^{key}\s*=.*$"
        if re.search(pattern, block):
            block = re.sub(pattern, lambda _: f"{key} = {encoded}", block)
        else:
            block = block.rstrip() + f"\n{key} = {encoded}\n\n"
    output = {path: text[:section.start()] + block + text[section.end():]}
    mirror = task / "metadata/task.json"
    raw = mirror.read_text(encoding="utf-8")
    data = json.loads(raw)
    updates["notes"] = calibration_notes(data.get("notes", ""))
    for key, value in updates.items():
        if key in data:
            raw, count = re.subn(rf'("{key}"\s*:\s*)"(?:\\.|[^"\\])*"',
                                lambda m: m[1] + json.dumps(value, ensure_ascii=False), raw)
            if count != 1:
                raise ValueError(f"重复元数据字段：{mirror}/{key}")
        else:
            end = raw.rfind("}")
            raw = raw[:end].rstrip() + ",\n  " + json.dumps(key) + ": " + json.dumps(value, ensure_ascii=False) + "\n" + raw[end:]
    parsed = json.loads(raw)
    if any(parsed.get(k) != v for k, v in updates.items()):
        raise ValueError(f"元数据镜像同步失败：{mirror}")
    output[mirror] = raw
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true", help="回填元数据；默认只校验")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    records = build_manifest(root)
    stream = io.StringIO()
    writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
    writer.writeheader()
    writer.writerows(records)
    planned = {root / MANIFEST: stream.getvalue()}
    for record in records:
        planned.update(metadata_updates(root, record))
    changes = {p: text for p, text in planned.items()
               if not p.exists() or p.read_text(encoding="utf-8") != text}
    if args.write:
        for path, content in changes.items():
            path.write_text(content, encoding="utf-8")
    elif changes:
        raise ValueError(f"难度未同步，共 {len(changes)} 个文件；首个：{next(iter(changes))}")
    print(f"难度核验通过：240 题；四模块各 easy 12 / medium 18 / hard 18 / expert 12；更新 {len(changes) if args.write else 0} 个文件。")


if __name__ == "__main__":
    main()
