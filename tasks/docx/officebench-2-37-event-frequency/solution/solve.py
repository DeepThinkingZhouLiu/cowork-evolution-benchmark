#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import subprocess
from collections import Counter
from email import policy
from email.parser import BytesParser
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import load_workbook


MODE = 'event_frequency'
OUTPUT_NAME = 'events.docx'
EXPECTED_HASHES = {'Alice.ics': '8559f5246583014288b3c81d98892874c7610c3586e95799035c6f77d571fd63'}


def digest(path: Path) -> str:
    import hashlib
    return hashlib.sha256(path.read_bytes()).hexdigest()


def setup(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Noto Sans CJK SC")
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(5)
    for name, size in (("Title", 20), ("Heading 1", 14), ("Heading 2", 11)):
        style = document.styles[name]
        style.font.name = "Liberation Sans"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Noto Sans CJK SC")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(31, 78, 121)


def title(document: Document, value: str, subtitle: str) -> None:
    paragraph = document.add_paragraph(style="Title")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run(value)
    sub = document.add_paragraph(subtitle)
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER


def table(document: Document, headers: list[str], rows: list[list[object]]) -> None:
    item = document.add_table(rows=1, cols=len(headers))
    item.style = "Table Grid"
    item.alignment = WD_TABLE_ALIGNMENT.CENTER
    for index, value in enumerate(headers):
        item.rows[0].cells[index].text = str(value)
    for values in rows:
        cells = item.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = str(value)


def parse_ics(path: Path) -> list[str]:
    return [value.strip() for value in re.findall(r"^SUMMARY:(.*)$", path.read_text(encoding="utf-8"), re.MULTILINE)]


def parse_email(path: Path) -> dict[str, str]:
    message = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
    part = message.get_body(preferencelist=("plain", "html"))
    raw = part.get_content().strip() if part else ""
    match = re.search(r"(?:^|\n)Sent at:\s*([^\r\n]+)", raw)
    if not match:
        raise ValueError(f"missing Sent at in {path.name}")
    body = raw[:match.start()].strip()
    return {"sent_at": match.group(1).strip(), "sender": str(message["From"] or ""), "subject": str(message["Subject"] or ""), "body": body}


def verify_inputs(input_dir: Path) -> None:
    actual = {path.name: digest(path) for path in input_dir.iterdir() if path.is_file()}
    if actual != EXPECTED_HASHES:
        raise ValueError("fixed OfficeBench input hash mismatch")


def build_event(input_dir: Path, output: Path) -> None:
    events = parse_ics(input_dir / "Alice.ics")
    counts = Counter(events)
    ranked = counts.most_common()
    document = Document()
    setup(document)
    title(document, "Event Frequency Analysis", "Source: Alice.ics")
    document.add_paragraph(f"Total events: {len(events)}")
    document.add_paragraph("Frequency by event: " + "; ".join(f"{name} {count}" for name, count in ranked))
    document.add_paragraph(f"Most frequent event: {ranked[0][0]}")
    document.add_paragraph(f"Count: {ranked[0][1]}")
    document.add_paragraph("The maximum is unique." if len(ranked) < 2 or ranked[0][1] > ranked[1][1] else "The maximum is tied.")
    document.save(output)


def build_midterm(input_dir: Path, output: Path) -> None:
    workbook = load_workbook(input_dir / "score.xlsx", data_only=True, read_only=True)
    rows = list(workbook.active.iter_rows(values_only=True))
    workbook.close()
    headers = [str(value).strip() for value in rows[0]]
    index = {value: position for position, value in enumerate(headers)}
    data = [row for row in rows[1:] if row[index["Name"]] is not None]
    m1 = sum(float(row[index["midterm1"]]) for row in data) / len(data)
    m2 = sum(float(row[index["midterm2"]]) for row in data) / len(data)
    document = Document()
    setup(document)
    title(document, "Midterm Score Report", "Computed from score.xlsx")
    table(document, ["Metric", "Result"], [["Students", len(data)], ["Midterm 1 average", f"{m1:.2f}"], ["Midterm 2 average", f"{m2:.2f}"], ["Midterm 2 - Midterm 1", f"{m2-m1:+.2f}"]])
    document.add_heading("Observation", level=1)
    direction = "lower" if m2 < m1 else "higher"
    document.add_paragraph(f"The Midterm 2 average is {abs(m2-m1):.2f} points {direction} than the Midterm 1 average. This is a descriptive cohort comparison, not a causal conclusion.")
    document.save(output)


def build_emails(input_dir: Path, output: Path) -> None:
    records = sorted((parse_email(path) for path in input_dir.glob("*.eml")), key=lambda item: item["sent_at"])
    document = Document()
    setup(document)
    title(document, "Bob's Emails in Chronological Order", "Four supplied messages")
    document.add_paragraph("The source messages do not contain Date headers. Chronology therefore uses each message's visible Sent at timestamp.")
    table(document, ["Sent at", "From", "Subject", "Body"], [[item["sent_at"], item["sender"], item["subject"], item["body"]] for item in records])
    document.save(output)


def build_paper(input_dir: Path, output: Path) -> None:
    source = subprocess.check_output(["pdftotext", str(input_dir / "paper.pdf"), "-"], text=True, errors="replace")
    for anchor in ("Attention Is All You Need", "28.4 BLEU", "41.0", "8 NVIDIA P100 GPUs"):
        if anchor not in source:
            raise ValueError(f"reviewed PDF anchor missing: {anchor}")
    document = Document()
    setup(document)
    title(document, "Attention Is All You Need", "Evidence-based reference summary of the supplied paper")

    document.add_heading("1. Paper information and purpose", level=1)
    document.add_paragraph("Vaswani et al., presented at the 31st Conference on Neural Information Processing Systems (NIPS 2017), address the sequential-computation bottleneck in recurrent sequence-transduction models. The paper evaluates whether an encoder-decoder can rely entirely on attention while improving translation quality and training parallelism.")
    table(document, ["Item", "Source-backed detail"], [
        ["Authors", "Ashish Vaswani; Noam Shazeer; Niki Parmar; Jakob Uszkoreit; Llion Jones; Aidan N. Gomez; Lukasz Kaiser; Illia Polosukhin"],
        ["Primary tasks", "WMT 2014 English-to-German and English-to-French machine translation"],
        ["Central proposal", "The Transformer, based solely on attention mechanisms, dispensing with recurrence and convolutions"],
    ])

    document.add_heading("2. Core contribution", level=1)
    document.add_paragraph("The Transformer retains the encoder-decoder structure but replaces recurrent and convolutional sequence processing with stacked self-attention and position-wise feed-forward layers. This permits substantially more parallel computation during training while allowing every position to draw information from distant positions through short paths.")

    document.add_heading("3. Architecture and mechanisms", level=1)
    table(document, ["Component", "Reviewed configuration", "Role"], [
        ["Encoder", "N = 6 identical layers", "Multi-head self-attention followed by a position-wise feed-forward network"],
        ["Decoder", "N = 6 identical layers", "Masked self-attention, encoder-decoder attention, and feed-forward network"],
        ["Representation", "d_model = 512; d_ff = 2048", "Shared hidden width and inner feed-forward width"],
        ["Multi-head attention", "h = 8; d_k = d_v = 64", "Attend to different representation subspaces in parallel"],
        ["Stabilization", "Residual connection around each sub-layer, followed by layer normalization", "Supports deep stacked computation"],
        ["Order", "Sinusoidal positional encoding", "Injects position because the model has no recurrence or convolution"],
    ])
    document.add_paragraph("Scaled dot-product attention computes compatibility from queries and keys, divides the dot products by the square root of d_k, applies softmax, and uses the resulting weights over values. The decoder mask prevents a position from attending to subsequent output positions.")

    document.add_heading("4. Why self-attention", level=1)
    document.add_paragraph("The paper compares three dimensions: per-layer complexity, minimum sequential operations, and maximum path length between positions. Full self-attention has O(n^2 d) per-layer complexity, O(1) sequential operations, and O(1) maximum path length; a recurrent layer requires O(n) sequential operations and an O(n) maximum path. The authors argue self-attention is faster than recurrence when sequence length n is smaller than representation dimension d, a common regime for the translation representations studied. They also observe that different heads can learn distinct syntactic or semantic patterns, suggesting interpretability benefits.")

    document.add_heading("5. Training and reported results", level=1)
    table(document, ["Evidence", "Reported value"], [
        ["English-German data", "About 4.5 million WMT 2014 sentence pairs; shared vocabulary of about 37,000 tokens"],
        ["English-French data", "36 million WMT 2014 sentences; 32,000 word-piece vocabulary"],
        ["Hardware", "One machine with 8 NVIDIA P100 GPUs"],
        ["Base training", "100,000 steps, about 12 hours"],
        ["Big training", "300,000 steps, 3.5 days"],
        ["WMT 2014 English-to-German", "28.4 BLEU, more than 2 BLEU above the best previously reported results including ensembles"],
        ["WMT 2014 English-to-French", "41.0 BLEU, a new single-model state of the art reported by the paper"],
    ])

    document.add_heading("6. Scope, limitations, and takeaway", level=1)
    document.add_paragraph("The empirical claims are demonstrated on the two machine-translation benchmarks in the supplied paper; they do not by themselves establish performance for every sequence or multimodal task. Full self-attention's O(n^2 d) complexity can be costly for very long sequences. The authors identify local or restricted attention for large inputs and outputs, extension beyond text to images, audio, and video, and making generation less sequential as future work.")
    document.add_paragraph("Takeaway: the paper shows that a carefully constructed all-attention encoder-decoder can achieve strong translation quality with greater training parallelism, establishing the architectural foundation later known as the Transformer.")
    document.save(output)


def solve(input_dir: Path, output_dir: Path) -> None:
    verify_inputs(input_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
    if MODE == "event_frequency":
        build_event(input_dir, output)
    elif MODE == "midterm_report":
        build_midterm(input_dir, output)
    elif MODE == "all_emails":
        build_emails(input_dir, output)
    elif MODE == "paper_summary":
        build_paper(input_dir, output)
    else:
        raise ValueError(MODE)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)
