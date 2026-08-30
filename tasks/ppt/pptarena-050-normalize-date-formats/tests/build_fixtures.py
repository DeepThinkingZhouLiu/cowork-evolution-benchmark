#!/usr/bin/env python3
"""只为 PPTArena Case 50 构造合法变体、日期/样式/保持负例与中等例。"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches


OUTPUT_NAME = "NormalizeDate_GroundTruth.pptx"


def output_path(root: Path, case: str) -> Path:
    target = root / case / OUTPUT_NAME
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


def shape_by_id(presentation: Presentation, slide_number: int, shape_id: int):
    matches = [shape for shape in presentation.slides[slide_number - 1].shapes if shape.shape_id == shape_id]
    if len(matches) != 1:
        raise ValueError(f"无法唯一定位：{slide_number}/{shape_id}")
    return matches[0]


def replace_in_shape(shape, old: str, new: str) -> None:
    total = 0
    for paragraph in shape.text_frame.paragraphs:
        runs = list(paragraph.runs)
        text = "".join(run.text for run in runs)
        index = text.find(old)
        if index < 0:
            continue
        if text.find(old, index + len(old)) >= 0:
            raise ValueError(f"负例替换目标不唯一：{old!r}")
        end = index + len(old)
        spans = []
        cursor = 0
        for run_index, run in enumerate(runs):
            next_cursor = cursor + len(run.text)
            spans.append((run_index, cursor, next_cursor))
            cursor = next_cursor
        start_info = next(item for item in spans if item[1] <= index < item[2])
        end_info = next(item for item in spans if item[1] < end <= item[2])
        start_run = runs[start_info[0]]
        end_run = runs[end_info[0]]
        start_offset = index - start_info[1]
        end_offset = end - end_info[1]
        if start_info[0] == end_info[0]:
            start_run.text = start_run.text[:start_offset] + new + start_run.text[end_offset:]
        else:
            start_run.text = start_run.text[:start_offset] + new
            for run_index in range(start_info[0] + 1, end_info[0]):
                runs[run_index].text = ""
            end_run.text = end_run.text[end_offset:]
        total += 1
    if total != 1:
        raise ValueError(f"负例替换计数异常：{old!r}，{total}")


def valid_alternate(golden: Path, output: Path) -> None:
    """只改变 run 分段，不改文本或逐字符样式，用于防止 Judge 硬编码 run 数量。"""
    presentation = Presentation(golden)
    shape = shape_by_id(presentation, 2, 7)
    paragraph = shape.text_frame.paragraphs[0]
    matches = [run for run in paragraph.runs if "February 25, 2022" in run.text]
    if len(matches) != 1:
        raise ValueError("合法变体的锁定 run 结构变化")
    run = matches[0]
    split = run.text.index("February 25, 2022") + len("February")
    suffix = run.text[split:]
    run.text = run.text[:split]
    added = paragraph.add_run()
    added.text = suffix
    added.font.language_id = run.font.language_id
    # python-pptx 默认把新 run 追加到段落末尾；本变体要求文字顺序完全不变。
    run._r.addnext(added._r)
    presentation.save(output)


def replace_case(golden: Path, output: Path, slide: int, shape_id: int, old: str, new: str) -> None:
    presentation = Presentation(golden)
    replace_in_shape(shape_by_id(presentation, slide, shape_id), old, new)
    presentation.save(output)


def wrong_style(golden: Path, output: Path) -> None:
    presentation = Presentation(golden)
    shape = shape_by_id(presentation, 9, 9)
    matches = [run for paragraph in shape.text_frame.paragraphs for run in paragraph.runs if "March 18, 2022" in run.text]
    if len(matches) != 1:
        raise ValueError("无法唯一定位 March 18 的 run")
    matches[0].font.italic = True
    presentation.save(output)


def delete_empty_placeholder(golden: Path, output: Path) -> None:
    presentation = Presentation(golden)
    shape = shape_by_id(presentation, 2, 4)
    if shape.text:
        raise ValueError("锁定对象不再是空占位符")
    shape._element.getparent().remove(shape._element)
    presentation.save(output)


def changed_geometry(golden: Path, output: Path) -> None:
    presentation = Presentation(golden)
    shape_by_id(presentation, 12, 9).left += Inches(0.50)
    presentation.save(output)


def changed_page_number(golden: Path, output: Path) -> None:
    presentation = Presentation(golden)
    shape_by_id(presentation, 7, 2).text = "70"
    presentation.save(output)


def remove_last_slide(golden: Path, output: Path) -> None:
    presentation = Presentation(golden)
    slide_id = presentation.slides._sldIdLst[-1]
    presentation.part.drop_rel(slide_id.rId)
    presentation.slides._sldIdLst.remove(slide_id)
    presentation.save(output)


def flatten_slide7(golden: Path, slide_image: Path, output: Path) -> None:
    presentation = Presentation(golden)
    slide = presentation.slides[6]
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)
    slide.shapes.add_picture(str(slide_image), 0, 0, width=presentation.slide_width, height=presentation.slide_height)
    presentation.save(output)


def hardcoded_marker(golden: Path, output: Path) -> None:
    presentation = Presentation(golden)
    box = presentation.slides[0].shapes.add_textbox(Inches(0.10), Inches(0.10), Inches(2.50), Inches(0.25))
    box.text = "solution/Golden"
    presentation.save(output)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--slide7-image", type=Path, required=True)
    parser.add_argument("--outdir", type=Path, required=True)
    args = parser.parse_args()

    (args.outdir / "empty").mkdir(parents=True, exist_ok=True)
    output_path(args.outdir, "corrupt").write_bytes(b"not a valid pptx\n")
    shutil.copy2(args.source, output_path(args.outdir, "unmodified"))
    valid_alternate(args.golden, output_path(args.outdir, "valid-alternate"))
    replace_case(args.golden, output_path(args.outdir, "missed-date"), 2, 7, "February 25, 2022", "25 February 2022")
    replace_case(args.golden, output_path(args.outdir, "wrong-year"), 8, 9, "Effective March 2, 2022 airlines", "Effective March 2, 2021 airlines")
    replace_case(
        args.golden,
        output_path(args.outdir, "wrong-numeric-interpretation"),
        7,
        9,
        "Further, on March 2, 2022, BIS",
        "Further, on February 3, 2022, BIS",
    )
    replace_case(args.golden, output_path(args.outdir, "missing-inferred-year"), 8, 9, "Effective March 2, 2022 airlines", "Effective March 2 airlines")
    replace_case(args.golden, output_path(args.outdir, "wrong-casing"), 3, 7, "march 5, 2022", "March 5, 2022")
    replace_case(args.golden, output_path(args.outdir, "changed-non-date-text"), 7, 8, "Us EXPORT CONTROLS", "Us EXPORT CONTROL")
    wrong_style(args.golden, output_path(args.outdir, "wrong-style"))
    delete_empty_placeholder(args.golden, output_path(args.outdir, "deleted-empty-placeholder"))
    changed_geometry(args.golden, output_path(args.outdir, "changed-geometry"))
    changed_page_number(args.golden, output_path(args.outdir, "changed-page-number"))
    remove_last_slide(args.golden, output_path(args.outdir, "removed-closing-slide"))
    flatten_slide7(args.golden, args.slide7_image, output_path(args.outdir, "flattened-slide7"))
    hardcoded_marker(args.golden, output_path(args.outdir, "hardcoded-marker"))
    extra = args.outdir / "extra-output-file"
    extra.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.golden, extra / OUTPUT_NAME)
    (extra / "debug.txt").write_text("debug\n", encoding="utf-8")


if __name__ == "__main__":
    main()
