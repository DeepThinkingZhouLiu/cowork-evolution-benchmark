#!/usr/bin/env python3
"""按冻结规则生成或校验 CoworkEvoBench 三段式 Release。"""

from __future__ import annotations

import argparse
import collections
import csv
import hashlib
import json
import random
import re
import subprocess
import tomllib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable


DEFAULT_RELEASE_ID = "cowork-evo-240-v0.5"
DOMAINS = ("ppt", "docx", "xlsx", "cross_office")
PARTITIONS = ("feedback", "selection", "final")
DOMAIN_TARGETS = {
    "ppt": {"feedback": 30, "selection": 10, "final": 20},
    "docx": {"feedback": 30, "selection": 10, "final": 20},
    "xlsx": {"feedback": 30, "selection": 10, "final": 20},
    "cross_office": {"feedback": 0, "selection": 0, "final": 60},
}
PPT_PILOT_TARGETS = {
    "feedback": {"easy": 6, "medium": 9, "hard": 9, "expert": 6},
    "selection": {"easy": 2, "medium": 3, "hard": 3, "expert": 2},
    "final": {"easy": 4, "medium": 6, "hard": 6, "expert": 4},
}
VISIBILITY = {
    "feedback": "detailed",
    "selection": "aggregate_only",
    "final": "sealed",
}
SPLIT_FILENAMES = {
    "feedback": "feedback.jsonl",
    "selection": "selection.jsonl",
    "final": "final.sealed.jsonl",
}


@dataclass
class Task:
    index: int
    domain: str
    slug: str
    root: Path
    metadata: dict
    input_hashes: set[str] = field(default_factory=set)
    content_sha256: str = ""
    source: str = ""
    effort_bucket: str = ""
    pilot_tier: str | None = None

    @property
    def instance_id(self) -> str:
        return f"{self.domain}/{self.slug}"


class DisjointSet:
    def __init__(self, size: int) -> None:
        self.parents = list(range(size))

    def find(self, value: int) -> int:
        parent = self.parents[value]
        if parent != value:
            self.parents[value] = self.find(parent)
        return self.parents[value]

    def union(self, left: int, right: int) -> None:
        left_root = self.find(left)
        right_root = self.find(right)
        if left_root != right_root:
            self.parents[right_root] = left_root


def _run_git(repo_root: Path, *args: str) -> str:
    completed = subprocess.run(
        ["git", *args],
        cwd=repo_root,
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _task_content_sha256(repo_root: Path, task_root: Path) -> str:
    relative = task_root.relative_to(repo_root).as_posix()
    raw = subprocess.run(
        ["git", "ls-files", "-z", "--", relative],
        cwd=repo_root,
        check=True,
        capture_output=True,
    ).stdout
    paths = sorted(item.decode("utf-8") for item in raw.split(b"\0") if item)
    if not paths:
        raise ValueError(f"任务目录没有已跟踪文件：{relative}")
    digest = hashlib.sha256()
    for item in paths:
        file_path = repo_root / item
        digest.update(item.encode("utf-8"))
        digest.update(b"\0")
        digest.update(bytes.fromhex(_sha256(file_path)))
        digest.update(b"\n")
    return digest.hexdigest()


def _canonical_source(provenance_id: str) -> str:
    source = provenance_id.split("/", 1)[0].strip().lower()
    if "officeval" in source or "omegause" in source:
        return "omegause-officeval"
    if "spreadsheetbench" in source:
        return "spreadsheetbench2"
    if source in {"spreadsheet-rl", "spreadsheet_rl"}:
        return "spreadsheet-rl"
    return source


def _load_ppt_pilot_tiers(repo_root: Path) -> dict[str, str]:
    path = repo_root / "benchmarks" / "ppt-difficulty-pilot-20260904.csv"
    with path.open(encoding="utf-8", newline="") as handle:
        return {
            row["task"]: row["pilot_tier"].strip().lower()
            for row in csv.DictReader(handle)
        }


def _discover_tasks(repo_root: Path) -> list[Task]:
    pilot_tiers = _load_ppt_pilot_tiers(repo_root)
    tasks: list[Task] = []
    for domain in DOMAINS:
        for task_toml in sorted((repo_root / "tasks" / domain).glob("*/task.toml")):
            payload = tomllib.loads(task_toml.read_text(encoding="utf-8"))
            metadata = payload["metadata"]
            task = Task(
                index=len(tasks),
                domain=domain,
                slug=task_toml.parent.name,
                root=task_toml.parent,
                metadata=metadata,
                source=_canonical_source(metadata["provenance_id"]),
                pilot_tier=pilot_tiers.get(task_toml.parent.name),
            )
            input_root = task.root / "environment" / "input_files"
            if not input_root.is_dir():
                input_root = task.root / "data" / "input_files"
            task.input_hashes = {
                _sha256(path) for path in input_root.rglob("*") if path.is_file()
            }
            task.content_sha256 = _task_content_sha256(repo_root, task.root)
            tasks.append(task)

    counts = collections.Counter(task.domain for task in tasks)
    if counts != collections.Counter({domain: 60 for domain in DOMAINS}):
        raise ValueError(f"任务数量不是四个模块各 60 道：{dict(counts)}")

    for domain in DOMAINS:
        domain_tasks = sorted(
            (task for task in tasks if task.domain == domain),
            key=lambda task: (int(task.metadata.get("estimated_minutes", 0)), task.instance_id),
        )
        for position, task in enumerate(domain_tasks):
            task.effort_bucket = f"Q{min(3, position * 4 // len(domain_tasks)) + 1}"

    ppt_tier_counts = collections.Counter(
        task.pilot_tier for task in tasks if task.domain == "ppt"
    )
    if ppt_tier_counts != collections.Counter(
        {"easy": 12, "medium": 18, "hard": 18, "expert": 12}
    ):
        raise ValueError(f"PPT 试校准档位不完整：{dict(ppt_tier_counts)}")
    return tasks


def _build_components(tasks: list[Task]) -> list[list[int]]:
    disjoint = DisjointSet(len(tasks))
    for field_name in ("leakage_group_id", "provenance_id"):
        groups: dict[str, list[int]] = collections.defaultdict(list)
        for task in tasks:
            groups[str(task.metadata[field_name])].append(task.index)
        for members in groups.values():
            for member in members[1:]:
                disjoint.union(members[0], member)

    hashes: dict[str, list[int]] = collections.defaultdict(list)
    for task in tasks:
        for input_hash in task.input_hashes:
            hashes[input_hash].append(task.index)
    for members in hashes.values():
        unique_members = sorted(set(members))
        for member in unique_members[1:]:
            disjoint.union(unique_members[0], member)

    components: dict[int, list[int]] = collections.defaultdict(list)
    for task in tasks:
        components[disjoint.find(task.index)].append(task.index)
    return sorted(
        (sorted(members) for members in components.values()),
        key=lambda members: tuple(tasks[index].instance_id for index in members),
    )


def _partition_score(
    domain: str,
    tasks: list[Task],
    assignments: dict[int, str],
) -> float:
    domain_tasks = [task for task in tasks if task.domain == domain]
    weights = {
        "source": 2.5,
        "task_family": 1.6,
        "dependency_scope": 1.2,
        "effort_bucket": 2.0,
    }
    if domain == "ppt":
        weights["pilot_tier"] = 20.0

    score = 0.0
    for attribute, weight in weights.items():
        def value(task: Task) -> str:
            if attribute == "source":
                return task.source
            if attribute == "pilot_tier":
                return str(task.pilot_tier)
            if attribute == "effort_bucket":
                return task.effort_bucket
            return str(task.metadata[attribute])

        totals = collections.Counter(value(task) for task in domain_tasks)
        actual = {
            partition: collections.Counter(
                value(task)
                for task in domain_tasks
                if assignments[task.index] == partition
            )
            for partition in PARTITIONS
        }
        for category, total in totals.items():
            for partition in PARTITIONS:
                target = total * DOMAIN_TARGETS[domain][partition] / 60
                score += (
                    weight
                    * (actual[partition][category] - target) ** 2
                    / (target + 0.5)
                )

    total_minutes = sum(int(task.metadata.get("estimated_minutes", 0)) for task in domain_tasks)
    for partition in PARTITIONS:
        actual_minutes = sum(
            int(task.metadata.get("estimated_minutes", 0))
            for task in domain_tasks
            if assignments[task.index] == partition
        )
        target_minutes = total_minutes * DOMAIN_TARGETS[domain][partition] / 60
        score += 0.25 * (actual_minutes - target_minutes) ** 2 / (target_minutes + 1)
    return score


def _assign_ppt(
    tasks: list[Task],
    components: list[list[int]],
    assignments: dict[int, str],
) -> None:
    units_by_tier: dict[str, list[list[int]]] = collections.defaultdict(list)
    for component in components:
        members = [index for index in component if tasks[index].domain == "ppt"]
        if not members:
            continue
        if len(component) != 1 or len(members) != 1:
            raise ValueError("PPT 任务出现未预期的跨题泄漏组件。")
        units_by_tier[str(tasks[members[0]].pilot_tier)].append(members)

    best: tuple[float, dict[int, str]] | None = None
    rng = random.Random(2026090601)
    for _ in range(4000):
        candidate = dict(assignments)
        for tier, units in units_by_tier.items():
            shuffled = units[:]
            rng.shuffle(shuffled)
            position = 0
            for partition in PARTITIONS:
                count = PPT_PILOT_TARGETS[partition][tier]
                for unit in shuffled[position : position + count]:
                    candidate[unit[0]] = partition
                position += count
        score = _partition_score("ppt", tasks, candidate)
        if best is None or score < best[0]:
            best = (score, candidate)
    if best is None:
        raise RuntimeError("无法生成 PPT 划分。")
    assignments.update({index: best[1][index] for index in best[1] if tasks[index].domain == "ppt"})


def _assign_domain(
    domain: str,
    tasks: list[Task],
    components: list[list[int]],
    assignments: dict[int, str],
) -> None:
    domain_indices = [task.index for task in tasks if task.domain == domain]
    fixed = {index: assignments[index] for index in domain_indices if index in assignments}
    remaining_units: list[list[int]] = []
    for component in components:
        members = [
            index
            for index in component
            if tasks[index].domain == domain and index not in fixed
        ]
        if not members:
            continue
        if len(members) != len(component):
            raise ValueError(f"{domain} 存在未固定的跨模块泄漏组件。")
        remaining_units.append(members)

    fixed_counts = collections.Counter(fixed.values())
    needed = {
        partition: DOMAIN_TARGETS[domain][partition] - fixed_counts[partition]
        for partition in PARTITIONS
    }
    if any(value < 0 for value in needed.values()):
        raise ValueError(f"{domain} 的强制分配超过目标容量：{needed}")

    rng = random.Random(2026090602 + DOMAINS.index(domain))
    best: tuple[float, dict[int, str]] | None = None
    for _ in range(25000):
        candidate = dict(fixed)
        capacities = dict(needed)
        units = remaining_units[:]
        rng.shuffle(units)
        units.sort(key=lambda unit: (-len(unit), rng.random()))
        feasible = True
        for unit in units:
            options = [
                partition
                for partition in PARTITIONS
                if capacities[partition] >= len(unit)
            ]
            if not options:
                feasible = False
                break
            weights = [max(0.01, capacities[partition] ** 1.3) for partition in options]
            partition = rng.choices(options, weights=weights, k=1)[0]
            for index in unit:
                candidate[index] = partition
            capacities[partition] -= len(unit)
        if not feasible or any(capacities.values()):
            continue
        score = _partition_score(domain, tasks, candidate)
        if best is None or score < best[0]:
            best = (score, candidate)
    if best is None:
        raise RuntimeError(f"无法为 {domain} 找到满足精确数量的分配。")
    assignments.update(best[1])


def _assign_tasks(tasks: list[Task], components: list[list[int]]) -> dict[int, str]:
    assignments: dict[int, str] = {}
    for component in components:
        domains = {tasks[index].domain for index in component}
        if "cross_office" in domains:
            partition = "final"
        elif len(domains) > 1:
            # 同一输入同时派生为 DOCX/XLSX 时统一留在 Feedback，避免污染验证和测试。
            partition = "feedback"
        else:
            continue
        for index in component:
            assignments[index] = partition

    _assign_ppt(tasks, components, assignments)
    _assign_domain("docx", tasks, components, assignments)
    _assign_domain("xlsx", tasks, components, assignments)
    return assignments


def _component_records(
    tasks: list[Task],
    components: list[list[int]],
    assignments: dict[int, str],
) -> tuple[list[dict], dict[int, str]]:
    input_hash_owners: dict[str, set[int]] = collections.defaultdict(set)
    for task in tasks:
        for input_hash in task.input_hashes:
            input_hash_owners[input_hash].add(task.index)

    records: list[dict] = []
    component_ids: dict[int, str] = {}
    for component in components:
        instance_ids = sorted(tasks[index].instance_id for index in component)
        partitions = {assignments[index] for index in component}
        if len(partitions) != 1:
            raise ValueError(f"泄漏组件跨分区：{instance_ids}")
        digest = hashlib.sha256("\n".join(instance_ids).encode("utf-8")).hexdigest()
        component_id = f"leakage-{digest[:16]}"
        for index in component:
            component_ids[index] = component_id
        shared_hashes = sorted(
            input_hash
            for index in component
            for input_hash in tasks[index].input_hashes
            if len(input_hash_owners[input_hash]) > 1
        )
        records.append(
            {
                "component_id": component_id,
                "partition": next(iter(partitions)),
                "instance_ids": instance_ids,
                "declared_leakage_group_ids": sorted(
                    {str(tasks[index].metadata["leakage_group_id"]) for index in component}
                ),
                "provenance_ids": sorted(
                    {str(tasks[index].metadata["provenance_id"]) for index in component}
                ),
                "shared_input_sha256": sorted(set(shared_hashes)),
            }
        )
    return records, component_ids


def _write_json(path: Path, payload: object) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _write_jsonl(path: Path, records: Iterable[dict]) -> None:
    path.write_text(
        "".join(json.dumps(record, ensure_ascii=False, sort_keys=True) + "\n" for record in records),
        encoding="utf-8",
    )


def _split_entry(task: Task, partition: str, component_id: str) -> dict:
    return {
        "instance_id": task.instance_id,
        "task_id": task.metadata["task_id"],
        "task_path": f"tasks/{task.instance_id}",
        "domain": task.domain,
        "evaluation_role": (
            "compositional_ood" if task.domain == "cross_office" else "in_domain"
        ),
        "task_family": task.metadata["task_family"],
        "source": task.source,
        "provenance_id": task.metadata["provenance_id"],
        "declared_leakage_group_id": task.metadata["leakage_group_id"],
        "leakage_component_id": component_id,
        "difficulty_level": task.metadata["difficulty_level"],
        "difficulty_tier": task.metadata["difficulty_tier"],
        "ppt_pilot_tier": task.pilot_tier,
        "structural_effort_stratum": task.effort_bucket,
        "content_sha256": task.content_sha256,
        "partition": partition,
    }


def _release_readme(entries: list[dict], release_id: str) -> str:
    counts = collections.Counter((entry["domain"], entry["partition"]) for entry in entries)
    lines = [
        f"# CoworkEvoBench 240 题三段式划分 {release_id.rsplit('-', 1)[-1]}",
        "",
        "本目录冻结 240 道 Harbor Task 的分区归属，不复制题目文件。`tasks/` 是唯一题目源；本目录只保存 Release、任务哈希、泄漏组件和 split 清单。",
        "",
        "| 模块         | Feedback / 训练反馈 | Selection / 验证 | Final / 测试 | 用途 |",
        "|--------------|--------------------:|-----------------:|-------------:|------|",
    ]
    purposes = {
        "ppt": "域内训练反馈、Champion 选择与最终测试",
        "docx": "域内训练反馈、Champion 选择与最终测试",
        "xlsx": "域内训练反馈、Champion 选择与最终测试",
        "cross_office": "Compositional OOD / Cross-Office Transfer",
    }
    padded_labels = {
        "ppt": "PPT         ",
        "docx": "DOCX        ",
        "xlsx": "XLSX        ",
        "cross_office": "Cross-Office",
    }
    for domain in DOMAINS:
        lines.append(
            f'| {padded_labels[domain]} | {counts[domain, "feedback"]:>19} | '
            f'{counts[domain, "selection"]:>16} | {counts[domain, "final"]:>12} | '
            f'{purposes[domain]} |'
        )
    lines.extend(
        [
            "| 合计         |                  90 |               30 |          120 | Final 包含 60 道域内题和 60 道组合式 OOD 题 |",
            "",
            "## 划分原则",
            "",
            "- Feedback 返回逐题轨迹、Rubric 得分和失败证据，用于 Harness Evolution。",
            "- Selection 只返回聚合指标，用于选择或回退 Candidate，不能用于逐题修补。",
            "- Final 在 Harness 锁定后一次性运行，用于论文最终报告。",
            "- PPT 按单轮试校准的 12/18/18/12 档位等比例分层；DOCX/XLSX 沿用按任务家族、来源、依赖范围和预计耗时得到的原分区。",
            "- 难度以本 Release 的逐题 `difficulty_*` 字段为准；回填难度不改变已冻结的任务分区。",
            "- 相同 `leakage_group_id`、相同 `provenance_id` 或共享任一输入 SHA256 的任务合并为同一泄漏组件，整个组件只能进入一个分区。",
            "- 与 Cross-Office 共享来源或输入的单模态任务一并进入 Final，避免 OOD 测试材料从 Feedback 或 Selection 泄漏。",
            "",
            "## 文件责任",
            "",
            "- `benchmark.json`：实验入口和三个分区的完整任务编号。",
            "- `dataset.lock`：240 道题的不可变内容哈希与 Git task-tree 标识。",
            "- `leakage-groups.jsonl`：泄漏组件、组成任务与分区。",
            "- `splits/*.jsonl`：逐题路径、来源、家族、难度、角色、哈希与分区。",
            f"- 已标注难度：{sum(entry['difficulty_tier'] in {'easy', 'medium', 'hard', 'expert'} for entry in entries)}/{len(entries)}；校准状态和来源见对应任务元数据。",
            "",
            "## 边界",
            "",
            "本仓库公开的是分区清单，因此 `sealed` 表示实验协议中的访问边界，而不是 Git 权限隔离。正式论文运行仍需由评测 Runner 在 Candidate 锁定前禁止读取 Final 的题目内容、Golden、Rubric 和 Judge。",
            "",
        ]
    )
    return "\n".join(lines)


def _load_frozen_assignments(release_root: Path) -> dict[str, str]:
    assignments: dict[str, str] = {}
    for partition in PARTITIONS:
        path = release_root / "splits" / SPLIT_FILENAMES[partition]
        if not path.is_file():
            raise FileNotFoundError(path)
        for entry in _read_jsonl(path):
            instance_id = str(entry["instance_id"])
            if instance_id in assignments:
                raise ValueError(f"来源 Release 任务重复：{instance_id}")
            assignments[instance_id] = partition
    return assignments


def _release_version(release_id: str) -> str:
    match = re.search(r"-v(\d+)\.(\d+)$", release_id)
    if match is None:
        raise ValueError(f"Release ID 必须以 -v<major>.<minor> 结尾：{release_id}")
    return f"{match.group(1)}.{match.group(2)}.0"


def generate_release(
    repo_root: Path,
    release_root: Path,
    release_id: str,
    source_release_root: Path | None = None,
) -> None:
    tasks = _discover_tasks(repo_root)
    components = _build_components(tasks)
    if source_release_root is None:
        assignments = _assign_tasks(tasks, components)
    else:
        frozen = _load_frozen_assignments(source_release_root)
        current_ids = {task.instance_id for task in tasks}
        if set(frozen) != current_ids:
            missing = sorted(current_ids - set(frozen))
            extra = sorted(set(frozen) - current_ids)
            raise ValueError(
                f"来源 Release 与当前任务集不一致：missing={missing}, extra={extra}"
            )
        assignments = {task.index: frozen[task.instance_id] for task in tasks}
    component_records, component_ids = _component_records(tasks, components, assignments)

    entries = sorted(
        (
            _split_entry(task, assignments[task.index], component_ids[task.index])
            for task in tasks
        ),
        key=lambda entry: entry["instance_id"],
    )
    release_root.mkdir(parents=True, exist_ok=True)
    split_root = release_root / "splits"
    split_root.mkdir(parents=True, exist_ok=True)
    for partition in PARTITIONS:
        _write_jsonl(
            split_root / SPLIT_FILENAMES[partition],
            (entry for entry in entries if entry["partition"] == partition),
        )
    _write_jsonl(release_root / "leakage-groups.jsonl", component_records)

    task_tree_oid = _run_git(repo_root, "rev-parse", "HEAD:tasks")
    dataset_lock = {
        "schema_version": "cowork-evo-bench/dataset-lock/v0.1",
        "release_id": release_id,
        "task_tree_git_oid": task_tree_oid,
        "content_hash_algorithm": "sha256(path_utf8 || NUL || sha256(file)_raw || LF)",
        "expected_task_count": 240,
        "tasks": [
            {
                "instance_id": entry["instance_id"],
                "task_id": entry["task_id"],
                "task_path": entry["task_path"],
                "content_sha256": entry["content_sha256"],
            }
            for entry in entries
        ],
    }
    _write_json(release_root / "dataset.lock", dataset_lock)
    dataset_lock_sha256 = _sha256(release_root / "dataset.lock")

    partitions = {
        partition: {
            "visibility": VISIBILITY[partition],
            "expected_count": sum(entry["partition"] == partition for entry in entries),
            "instance_ids": [
                entry["instance_id"]
                for entry in entries
                if entry["partition"] == partition
            ],
        }
        for partition in PARTITIONS
    }
    benchmark = {
        "api_version": "cowork-evo-bench/v0.1",
        "kind": "Benchmark",
        "metadata": {
            "id": release_id,
            "name": f"CoworkEvoBench 240 题三段式划分 {release_id.rsplit('-', 1)[-1]}",
            "version": _release_version(release_id),
            "description": "PPT/DOCX/XLSX 各 30/10/20，Cross-Office 60 道全部作为组合式 OOD Final。",
        },
        "spec": {
            "task_source": {
                "registry": "git+https://github.com/DeepThinkingZhouLiu/cowork-evolution-benchmark#tasks",
                "revision": f"git-tree:{task_tree_oid}",
                "manifest_sha256": dataset_lock_sha256,
            },
            "evaluator": {
                "adapter": "harbor-office",
                "version": "0.21.0",
                "result_schema": "cowork-evo-bench/solver-result/v0.1",
            },
            "partitions": partitions,
            "expected_total": 240,
        },
    }
    _write_json(release_root / "benchmark.json", benchmark)
    (release_root / "README.md").write_text(
        _release_readme(entries, release_id), encoding="utf-8"
    )


def _read_jsonl(path: Path) -> list[dict]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line]


def validate_release(repo_root: Path, release_root: Path) -> None:
    benchmark = json.loads((release_root / "benchmark.json").read_text(encoding="utf-8"))
    dataset_lock = json.loads((release_root / "dataset.lock").read_text(encoding="utf-8"))
    if _sha256(release_root / "dataset.lock") != benchmark["spec"]["task_source"]["manifest_sha256"]:
        raise ValueError("benchmark.json 中的 manifest_sha256 与 dataset.lock 不一致。")

    task_tree_oid = _run_git(repo_root, "rev-parse", "HEAD:tasks")
    if dataset_lock["task_tree_git_oid"] != task_tree_oid:
        raise ValueError("当前 tasks Git tree 与 dataset.lock 不一致。")

    entries: list[dict] = []
    for partition in PARTITIONS:
        split_entries = _read_jsonl(release_root / "splits" / SPLIT_FILENAMES[partition])
        expected_ids = benchmark["spec"]["partitions"][partition]["instance_ids"]
        actual_ids = [entry["instance_id"] for entry in split_entries]
        if actual_ids != expected_ids:
            raise ValueError(f"{partition} 的 JSONL 与 benchmark.json 不一致。")
        entries.extend(split_entries)

    if len(entries) != 240 or len({entry["instance_id"] for entry in entries}) != 240:
        raise ValueError("分区没有恰好覆盖 240 道互异任务。")
    counts = collections.Counter((entry["domain"], entry["partition"]) for entry in entries)
    for domain, targets in DOMAIN_TARGETS.items():
        for partition, expected_count in targets.items():
            if counts[domain, partition] != expected_count:
                raise ValueError(
                    f"{domain}/{partition} 应为 {expected_count}，实际为 {counts[domain, partition]}。"
                )

    locked = {entry["instance_id"]: entry for entry in dataset_lock["tasks"]}
    for entry in entries:
        task_root = repo_root / entry["task_path"]
        if not (task_root / "task.toml").is_file():
            raise ValueError(f"任务路径不存在：{entry['task_path']}")
        actual_hash = _task_content_sha256(repo_root, task_root)
        if actual_hash != entry["content_sha256"] or actual_hash != locked[entry["instance_id"]]["content_sha256"]:
            raise ValueError(f"任务内容哈希不一致：{entry['instance_id']}")

    entry_partitions = {entry["instance_id"]: entry["partition"] for entry in entries}
    components = _read_jsonl(release_root / "leakage-groups.jsonl")
    covered: list[str] = []
    for component in components:
        partitions = {entry_partitions[instance_id] for instance_id in component["instance_ids"]}
        if partitions != {component["partition"]}:
            raise ValueError(f"泄漏组件跨分区：{component['component_id']}")
        covered.extend(component["instance_ids"])
    if sorted(covered) != sorted(entry_partitions):
        raise ValueError("leakage-groups.jsonl 没有恰好覆盖 240 道题。")

    ppt_counts = collections.Counter(
        (entry["partition"], entry["ppt_pilot_tier"])
        for entry in entries
        if entry["domain"] == "ppt"
    )
    for partition, targets in PPT_PILOT_TARGETS.items():
        for tier, expected_count in targets.items():
            if ppt_counts[partition, tier] != expected_count:
                raise ValueError(f"PPT {partition}/{tier} 分层数量不一致。")
    print("Release 校验通过：240 道任务、90/30/120 分区、0 个跨分区泄漏组件。")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--release-id", default=DEFAULT_RELEASE_ID)
    parser.add_argument(
        "--source-release",
        help="生成新版本时复用指定 Release 的分区归属，避免任务内容变更导致重新洗牌。",
    )
    parser.add_argument("--check", action="store_true", help="只校验已生成的 Release")
    args = parser.parse_args()
    repo_root = args.repo_root.resolve()
    release_root = repo_root / "benchmarks" / args.release_id
    if not args.check:
        source_release_root = (
            repo_root / "benchmarks" / args.source_release
            if args.source_release
            else None
        )
        generate_release(
            repo_root,
            release_root,
            args.release_id,
            source_release_root=source_release_root,
        )
    validate_release(repo_root, release_root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
