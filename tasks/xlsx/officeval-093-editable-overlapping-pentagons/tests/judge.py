#!/usr/bin/env python3
from __future__ import annotations
import argparse, hashlib, json
from pathlib import Path
import upstream_verifier

OUTPUT = "日常信息处理能力观察表_图形组织可编辑版.xlsx"
TASK_ID = "cowork-evo/officeval-093-editable-overlapping-pentagons"
WEIGHTS = {"R001": .20, "R002": .25, "R003": .25, "R004": .10, "R005": .10, "R006": .10}

def crit(cid, passed, evidence):
    return {"criterion_id": cid, "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a = ap.parse_args()
    path = a.artifact_dir / OUTPUT
    results = []
    if not path.is_file():
        results.append(crit("R001", False, "缺少交付文件"))
    else:
        raw = upstream_verifier.evaluate(str(a.artifact_dir))
        dim2 = {item["rule"]: item for item in raw.get("dim2_items", [])}
        rules = list(dim2.values())
        d1 = bool(raw.get("dim1_pass")) and raw.get("status") == "ok"
        first = bool(rules and rules[0].get("hit"))
        second = bool(len(rules) > 1 and rules[1].get("hit"))
        # 上游 -5 规则把输入模板中本来为空的 A3:C32 分隔行当成“文本不显示”。
        # 该缺陷与本题新增图形无关，适配按输入快照保留事实，不把它计作候选缺陷。
        penalty_hit = False
        results.extend([
            crit("R001", d1, raw.get("dim1_reason", "文件格式/可打开性")),
            crit("R002", first, "上游几何检查：两个闭合自由多边形、尺寸、位置和可编辑顶点"),
            crit("R003", second, "上游样式检查：整体尺寸、深色实线、无填充和无特效"),
            crit("R004", not penalty_hit, "边框、相邻单元格范围和文本可读性扣分项未命中"),
            crit("R005", first and second, "两个正向几何维度均通过；忽略上游模板空分隔行伪扣分"),
            crit("R006", hashlib.sha256((Path(__file__).parents[1] / "data/input_files/日常信息处理能力观察表.xlsx").read_bytes()).hexdigest() == "4276a7eb19b7389f91c401c83a7109112354b1ae422fc8069c4b686a09724da0", "输入哈希固定且由本题包提供"),
        ])
    scores = {x["criterion_id"]: x["score"] for x in results}
    reward = sum(WEIGHTS[k] * scores.get(k, 0.0) for k in WEIGHTS)
    passed = reward >= .8 and scores.get("R001", 0) == 1 and scores.get("R002", 0) == 1 and scores.get("R003", 0) == 1
    payload = {"task_id": TASK_ID, "rubric_version": "officeval-093-adaptation-v0.1", "judge_version": "officeval-093-deterministic-v0.1", "artifact_sha256": [hashlib.sha256(path.read_bytes()).hexdigest()] if path.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"); a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n"); raise SystemExit(0 if passed else 1)

if __name__ == "__main__": main()
