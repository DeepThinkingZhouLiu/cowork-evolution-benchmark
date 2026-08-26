#!/usr/bin/env python3
"""officeval_010 的独立 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

import upstream_verifier


TASK_ID = "cowork-evo/officeval-010-editable-technical-roadmap-reconstruction"
OUTPUT_NAME = "技术路线图_删除第四行_可编辑版.docx"
RUBRIC_VERSION = "omegause-officeval-010-adaptation-v0.1"
JUDGE_VERSION = "officeval-010-adaptation-v0.1"
REQUIRED = ["技术路线图", "研究导入：明确目标与范围", "政策信息梳理", "资料归集校核", "参照案例对照", "业务流程梳理", "需求要点识别与目标分解", "可行性确认", "完善资料与方案", "运行环境综合研判", "外部条件研判", "内部能力评阅", "政策环境扫描", "需求场景识别", "协同路径分析", "资源储备评估", "流程执行能力", "技术支撑能力", "运营协调机制", "EFE评阅清单", "适配矩阵", "IFE评阅清单", "实施路径定位与方案设计", "组织保障与质量闭环", "归纳与展望", "导入", "核验", "研判", "形成"]
REMOVED = ["资源能力盘点", "风险问题排查", "合规性初步审查"]
SPECS = {
    "R001": ("指定 DOCX 可打开且为一页 A4", 3, "hurdle"),
    "R002": ("指定第四行三个检查框已删除，其余 29 个节点文字完整", 7, "hurdle"),
    "R003": ("标题、分隔线、主流程矩形和文字样式符合来源规格", 5, "positive"),
    "R004": ("导入到可行性确认及回路的上半段连接关系完整", 6, "positive"),
    "R005": ("研判分支、七个竖框及 EFE/IFE 适配关系完整", 6, "positive"),
    "R006": ("形成阶段三节点与四个阶段标签完整", 4, "positive"),
    "R007": ("页面不含图片，节点和连接线均为原生可编辑 VML 对象", 6, "hurdle"),
    "R008": ("LibreOffice 实际渲染为一页且关键节点可提取", 3, "positive"),
    "R009": ("若原图仍被嵌入、节点缺失或被删除行仍存在，扣 10 分", -10, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {"criterion_id": identifier, "description": description, "weight": weight, "type": kind, "passed": bool(passed), "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def render_check(path: Path) -> tuple[bool, str]:
    if not shutil.which("libreoffice") or not shutil.which("pdfinfo") or not shutil.which("pdftotext"):
        return False, "缺少渲染依赖"
    with tempfile.TemporaryDirectory(prefix="officeval010-render-") as tmp:
        proc = subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(path)], capture_output=True, text=True, timeout=240)
        pdf = Path(tmp) / f"{path.stem}.pdf"
        if proc.returncode != 0 or not pdf.is_file():
            return False, f"渲染失败：{proc.stderr[-300:]}"
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=30).stdout
        pages = next((int(line.split(":", 1)[1]) for line in info.splitlines() if line.startswith("Pages:")), 0)
        text = subprocess.run(["pdftotext", str(pdf), "-"], capture_output=True, text=True, timeout=60).stdout.replace("\n", "")
        ok = pages == 1 and all(value in text for value in ["技术路线图", "可行性确认", "归纳与展望"])
        return ok, f"页数={pages}，标题/菱形/末节点可提取={ok}"


def evaluate(artifact_dir: Path) -> dict:
    output = artifact_dir / OUTPUT_NAME
    results = []
    try:
        model = upstream_verifier.DocxRoadmap(output)
        loaded = model.load()
        results.append(criterion("R001", loaded and model.section_count == 1, f"loaded={loaded}，节数={model.section_count}，页面={model.page_w_pt:.1f}×{model.page_h_pt:.1f}pt"))
        if not loaded:
            return build_result(results, output)
    except Exception as exc:
        results.append(criterion("R001", False, f"缺失或损坏：{exc}"))
        return build_result(results, output)

    texts = [shape.text.replace("\n", "") for shape in model.shapes]
    missing = [value for value in REQUIRED if value not in texts]
    present_removed = [value for value in REMOVED if value in texts]
    results.append(criterion("R002", not missing and not present_removed, f"缺失={missing}，删除行残留={present_removed}"))

    upstream = upstream_verifier.evaluate(str(artifact_dir))
    items = upstream.get("dim2_items", [])
    hits = [bool(item.get("hit")) for item in items]
    # 上游共 32 条确定性规则；按路线图阶段聚合，不用 Golden 哈希。
    results.append(criterion("R003", len(hits) >= 5 and sum(hits[:5]) >= 4, f"标题/基础样式={hits[:5]}"))
    results.append(criterion("R004", len(hits) >= 22 and sum(hits[5:22]) >= 13, f"上半段命中={sum(hits[5:22])}/17"))
    scorer = upstream_verifier.Scorer(model)
    row7_ok = scorer._check_seventh_row()
    row8_ok = scorer._check_eighth_row(model.find_shape_text("EFE评阅清单"), model.find_shape_text("适配矩阵"), model.find_shape_text("IFE评阅清单"))
    diagonal_ok = scorer._check_diagonal_arrows()
    branch_line_ids = {line.id for line in model.lines}
    branch_ids = {"external-bus", "external-child-1", "external-child-2", "external-child-3", "internal-bus", "internal-child-1", "internal-child-2", "internal-child-3", "internal-child-4"}
    analysis_ok = row7_ok and row8_ok and diagonal_ok and branch_ids.issubset(branch_line_ids)
    results.append(criterion("R005", analysis_ok, f"七竖框={row7_ok}，第八行={row8_ok}，斜箭头={diagonal_ok}，分支线完整={branch_ids.issubset(branch_line_ids)}；上游分支候选包含左侧阶段标签的已知误判={sum(hits[22:29]) if len(hits) >= 29 else 0}"))

    formation_nodes = all(model.find_shape_text(text) is not None for text in ["实施路径定位与方案设计", "组织保障与质量闭环", "归纳与展望", "导入", "核验", "研判", "形成"])
    formation_lines = {"matrix-path", "path-org", "org-summary"}.issubset(branch_line_ids)
    stage_ok = scorer._check_stage_boxes()
    results.append(criterion("R006", formation_nodes and formation_lines and stage_ok, f"形成节点={formation_nodes}，三段箭头={formation_lines}，阶段标签={stage_ok}"))

    with zipfile.ZipFile(output) as archive:
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        xml = etree.fromstring(archive.read("word/document.xml"))
    ns = {"v": "urn:schemas-microsoft-com:vml"}
    native_shapes = len(xml.xpath(".//v:rect | .//v:roundrect | .//v:shape", namespaces=ns))
    native_lines = len(xml.xpath(".//v:line", namespaces=ns))
    editable = not media and native_shapes >= 30 and native_lines >= 25
    results.append(criterion("R007", editable, f"媒体={len(media)}，原生形状={native_shapes}，原生连线={native_lines}"))
    render_pass, render_evidence = render_check(output)
    results.append(criterion("R008", render_pass, render_evidence))
    severe = bool(media) or bool(missing) or bool(present_removed)
    results.append(criterion("R009", severe, f"严重错误={severe}，上游={upstream.get('total_score')}/{upstream.get('max_score')}"))
    return build_result(results, output)


def build_result(results: list[dict], output: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"]) + sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, earned / total))
    hurdles = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdles:
        reward = 0.0
    normalized = [{"criterion_id": item["criterion_id"], "score": (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"], "evidence": item["evidence"], "method": item["method"], "raw": {"description": item["description"], "weight": item["weight"], "type": item["type"], "passed": item["passed"]}} for item in results]
    digest = hashlib.sha256(output.read_bytes() if output.is_file() else b"").hexdigest()
    return {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [digest], "criterion_results": normalized, "reward": round(reward, 6), "passed": reward >= 0.8, "hurdle_failures": hurdles}


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
