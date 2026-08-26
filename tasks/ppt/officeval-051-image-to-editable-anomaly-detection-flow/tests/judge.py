#!/usr/bin/env python3
"""officeval_051 的独立 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


OUTPUT_NAME = "anomaly_detection_editable_flow.pptx"
TASK_ID = "cowork-evo/officeval-051-image-to-editable-anomaly-detection-flow"
RUBRIC_VERSION = "omegause-officeval-051-adaptation-v0.1"
JUDGE_VERSION = "officeval-051-adaptation-v0.1"
EMU_PER_CM = 360000.0
SPECS = {
    "R001": ("指定 PPTX 可打开，严格为 1 页宽屏", 4, "hurdle"),
    "R002": ("八个主要区域标题与核心节点文本完整", 7, "hurdle"),
    "R003": ("多变量时序卡含网格和四条不同颜色折线", 4, "positive"),
    "R004": ("序列编码器与模式编码器的节点和 5×3 特征矩阵完整", 5, "positive"),
    "R005": ("融合与表示学习的四阶段、特征图标与三条向右连接线完整", 6, "positive"),
    "R006": ("异常评分和决策区的矢量、分布曲线、阈值与输出完整", 5, "positive"),
    "R007": ("训练目标含两个损失、加号、分隔线和总损失公式", 4, "positive"),
    "R008": ("数据流、辅助流和图例内容完整", 4, "positive"),
    "R009": ("页面不含位图且由至少 140 个原生可编辑对象组成", 6, "hurdle"),
    "R010": ("页面对象不越界，主要区域无大面积重叠", 3, "positive"),
    "R011": ("LibreOffice 可实际渲染为 1 页，关键文本可读", 4, "positive"),
    "R012": ("若嵌入位图、缺少核心区域或不是单页，扣 15 分", -15, "penalty"),
}
REQUIRED = [
    "Multivariate Time Series", "1. Sequence Encoder", "Gated Recurrent Unit", "2. Pattern Encoder", "Temporal Embedding",
    "3. Fusion & Representation Learning", "Concatenate", "Self-Attention Block", "Feed-Forward Network", "Layer Normalization",
    "4. Anomaly Scoring", "Reconstruction Module", "Anomaly Score", "5. Decision", "Thresholding", "Normal / Anomaly",
    "6. Training Objective", "Reconstruction Loss", "Lrec", "Regularization Loss", "Lreg", "Total Loss: L = Lrec + λLreg",
    "Legend:", "Data Flow", "Auxiliary Flow", "Vector / Feature",
]


def cm(value: int) -> float:
    return value / EMU_PER_CM


def norm(value: str) -> str:
    return re.sub(r"\s+", "", value).lower()


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {"criterion_id": identifier, "score": 1.0 if passed else 0.0, "evidence": evidence,
            "method": "deterministic", "raw": {"description": description, "weight": weight, "type": kind, "passed": bool(passed)}}


def render(path: Path) -> tuple[bool, str]:
    with tempfile.TemporaryDirectory(prefix="officeval051-judge-") as temp_dir:
        completed = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, str(path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=120)
        pdf = Path(temp_dir) / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf.is_file():
            return False, f"转 PDF 失败：{completed.stdout[-300:]}"
        info = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, text=True, timeout=30).stdout
        pages_match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE); pages = int(pages_match.group(1)) if pages_match else 0
        txt = Path(temp_dir) / "out.txt"; subprocess.run(["pdftotext", str(pdf), str(txt)], check=True, timeout=30)
        rendered = norm(txt.read_text(errors="replace"))
        keys = ["Sequence Encoder", "Fusion & Representation Learning", "Anomaly Score", "Total Loss"]
        ok = pages == 1 and all(norm(key) in rendered for key in keys)
        return ok, f"页数={pages}，关键文本可读={ok}"


def count(slide, name: str, *, starts: bool = False) -> int:
    return sum(1 for item in slide.shapes if item.name.startswith(name) if starts) if starts else sum(1 for item in slide.shapes if item.name == name)


def evaluate(path: Path) -> dict:
    criteria = []
    try:
        prs = Presentation(path)
    except Exception as exc:
        for identifier in SPECS:
            criteria.append(criterion(identifier, False, f"打开失败：{exc}"))
        return make_result(path, criteria)
    size_ok = abs(cm(prs.slide_width) - 38.4) <= .05 and abs(cm(prs.slide_height) - 20.0) <= .05
    criteria.append(criterion("R001", len(prs.slides) == 1 and size_ok, f"页数={len(prs.slides)}，页面={cm(prs.slide_width):.2f}×{cm(prs.slide_height):.2f}cm"))
    slide = prs.slides[0] if prs.slides else None
    content = "\n".join(item.text for item in slide.shapes if getattr(item, "has_text_frame", False)) if slide else ""
    missing = [value for value in REQUIRED if norm(value) not in norm(content)]
    criteria.append(criterion("R002", not missing, f"缺失={missing}"))

    grid = count(slide, "timeseries-grid") if slide else 0; series = count(slide, "timeseries-line") if slide else 0
    criteria.append(criterion("R003", grid >= 6 and series >= 24, f"网格={grid}，折线段={series}"))
    gru = count(slide, "gru-", starts=True) if slide else 0; cells = count(slide, "embedding-cell") if slide else 0
    criteria.append(criterion("R004", gru == 2 and cells == 15, f"GRU={gru}，特征单元={cells}"))
    fusion_cards = sum(1 for item in slide.shapes if re.fullmatch(r"fusion-card-\d+", item.name)) if slide else 0
    concat = count(slide, "concat-cell") if slide else 0; attention = count(slide, "attention-node") if slide else 0
    network = count(slide, "network-node") if slide else 0; norm_curve = count(slide, "normalization-curve") if slide else 0
    arrows = count(slide, "fusion-arrow") if slide else 0
    criteria.append(criterion("R005", (fusion_cards, concat, attention, network, norm_curve, arrows) == (4, 8, 4, 6, 6, 3),
                               f"阶段/矩阵/注意力/网络/曲线/箭头={fusion_cards}/{concat}/{attention}/{network}/{norm_curve}/{arrows}"))
    vectors = count(slide, "reconstruction-vector") if slide else 0; bell = count(slide, "anomaly-bell-curve") if slide else 0
    scoring_ok = all(count(slide, name) == 1 for name in ["scoring-frame", "reconstruction-module", "anomaly-score", "decision-frame", "thresholding", "normal-anomaly"]) if slide else False
    criteria.append(criterion("R006", scoring_ok and vectors == 4 and bell == 7, f"结构={scoring_ok}，特征块={vectors}，分布曲线段={bell}"))
    training_ok = slide and all(count(slide, name) == 1 for name in ["training-frame", "reconstruction-loss", "regularization-loss", "loss-separator"])
    criteria.append(criterion("R007", bool(training_ok), f"训练目标结构={bool(training_ok)}"))
    flows = sum(1 for item in slide.shapes if any(item.name.startswith(prefix) for prefix in ["input-to-", "sequence-to-", "pattern-to-", "fusion-to-", "training-to-"])) if slide else 0
    legend_ok = slide and all(count(slide, name) == 1 for name in ["legend-frame", "legend-data-arrow", "legend-aux-arrow", "legend-vector"])
    criteria.append(criterion("R008", flows >= 15 and bool(legend_ok), f"主/辅助流段={flows}，图例={bool(legend_ok)}"))

    pictures = sum(1 for item in slide.shapes if item.shape_type == MSO_SHAPE_TYPE.PICTURE) if slide else 0
    natives = len(slide.shapes) - pictures if slide else 0
    criteria.append(criterion("R009", pictures == 0 and natives >= 140, f"图片={pictures}，原生对象={natives}"))
    inside = slide and all(item.left >= -1000 and item.top >= -1000 and item.left + item.width <= prs.slide_width + 1000 and item.top + item.height <= prs.slide_height + 1000 for item in slide.shapes)
    criteria.append(criterion("R010", bool(inside), f"全部对象在页内={bool(inside)}"))
    render_ok, render_evidence = render(path); criteria.append(criterion("R011", render_ok, render_evidence))
    severe = len(prs.slides) != 1 or pictures > 0 or bool(missing)
    criteria.append(criterion("R012", not severe, f"严重错误={severe}"))
    return make_result(path, criteria)


def make_result(path: Path, criteria: list[dict]) -> dict:
    hurdles = [item["criterion_id"] for item in criteria if item["raw"]["type"] == "hurdle" and not item["raw"]["passed"]]
    total = sum(item["raw"]["weight"] for item in criteria if item["raw"]["weight"] > 0)
    earned = sum(item["raw"]["weight"] for item in criteria if item["raw"]["weight"] > 0 and item["raw"]["passed"])
    penalties = sum(abs(item["raw"]["weight"]) for item in criteria if item["raw"]["weight"] < 0 and not item["raw"]["passed"])
    reward = 0.0 if hurdles else max(0.0, min(1.0, (earned - penalties) / total))
    return {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION,
            "artifact_sha256": [hashlib.sha256(path.read_bytes()).hexdigest()] if path.is_file() else [],
            "criterion_results": criteria, "reward": reward, "passed": reward >= .8 and not hurdles, "hurdle_failures": hurdles}


def main() -> None:
    parser = argparse.ArgumentParser(); parser.add_argument("--artifact-dir", type=Path, required=True); parser.add_argument("--result", type=Path, required=True); parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args(); output = args.artifact_dir / OUTPUT_NAME
    evaluation = evaluate(output) if output.is_file() else {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [], "criterion_results": [], "reward": 0.0, "passed": False, "hurdle_failures": ["missing_output"]}
    args.result.parent.mkdir(parents=True, exist_ok=True); args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(evaluation, ensure_ascii=False, indent=2), encoding="utf-8"); args.reward_file.write_text(f"{evaluation['reward']:.6f}\n", encoding="utf-8")
    print(json.dumps(evaluation, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
