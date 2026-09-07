from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document


WEIGHTS = {"R001": .20, "R002": .20, "R003": .16, "R004": .14, "R005": .15, "R006": .10, "R007": .05}
SOURCE_SHA = "1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf"


def criterion(cid: str, passed: bool, evidence: str) -> dict:
    return {"criterion_id": cid, "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def valid_doc(path: Path) -> bool:
    return path.is_file() and zipfile.is_zipfile(path)


def paragraphs(path: Path) -> list[str]:
    return [p.text for p in Document(path).paragraphs]


def render_page_count(path: Path) -> tuple[bool, str]:
    with tempfile.TemporaryDirectory(prefix="officebench-3-22-") as temp:
        result = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", temp, str(path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=120)
        pdf = Path(temp) / f"{path.stem}.pdf"
        if result.returncode != 0 or not pdf.is_file():
            return False, result.stdout[-300:]
        info = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, text=True, timeout=30).stdout
        match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
        return True, f"页数={int(match.group(1)) if match else 0}"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    source = Path(__file__).parents[1] / "data/input_files/homework.jpg"
    first = args.artifact_dir / "homework.docx"
    second = args.artifact_dir / "homework/homework.docx"
    results: list[dict] = []
    try:
        first_ok = valid_doc(first)
        second_ok = valid_doc(second)
        first_text = paragraphs(first) if first_ok else []
        second_text = paragraphs(second) if second_ok else []
        nonempty = bool(first_text and any(x.strip() for x in first_text))
        compact = re.sub(r"\s+", " ", "\n".join(first_text)).casefold()
        header_hits = sum(x in compact for x in ("cse 240a", "homework 2", "pipelining", "memory", "ryan wang", "a012345", "june 1, 2021"))
        question_hits = sum(x in compact for x in ("pipeline", "branch", "cache", "memory", "cycle", "instruction"))
        results.append(criterion("R001", first_ok and second_ok, f"根文件可打开={first_ok}，子目录文件可打开={second_ok}"))
        results.append(criterion("R002", header_hits >= 6 and question_hits >= 4 and len(compact) >= 500, f"作业锚点={header_hits}/7，题目主题={question_hits}/6，字符={len(compact)}"))
        results.append(criterion("R003", first_text == second_text and nonempty, f"两份 OCR 文本一致={first_text == second_text}"))
        results.append(criterion("R004", len(first_text) >= 1 and len(second_text) >= 1, "两个 DOCX 均保留可编辑段落结构"))
        results.append(criterion("R005", all("TODO" not in x and "placeholder" not in x.lower() for x in first_text + second_text), "无占位内容"))
        digest = hashlib.sha256(source.read_bytes()).hexdigest() if source.is_file() else "missing"
        results.append(criterion("R006", digest == SOURCE_SHA, f"输入 SHA256={digest}"))
        render_ok, render_evidence = render_page_count(first) if first_ok else (False, "根文件不可渲染")
        results.append(criterion("R007", render_ok, render_evidence))
    except Exception as exc:
        results = [criterion("R001", False, f"读取失败: {exc}")]
    values = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[key] * values.get(key, 0.0) for key in WEIGHTS)
    passed = reward >= 0.8 and values.get("R001", 0.0) == 1.0 and values.get("R002", 0.0) == 1.0
    payload = {
        "task_id": "cowork-evo/officebench-3-22-homework-folder",
        "rubric_version": "officebench-3-22-homework-folder-v0.3",
        "judge_version": "officebench-3-22-homework-folder-deterministic-v0.3",
        "artifact_sha256": [hashlib.sha256(path.read_bytes()).hexdigest() for path in (first, second) if path.is_file()],
        "criterion_results": results,
        "reward": reward,
        "passed": passed,
    }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
