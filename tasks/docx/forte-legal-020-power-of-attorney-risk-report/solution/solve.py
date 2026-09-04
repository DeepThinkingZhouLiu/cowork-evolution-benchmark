#!/usr/bin/env python3
"""FORTE Legal-020 的逐题参考解法。

六份结论的等级和风险项来自上游八条真值 Rubric；触发证据、风险说明与
调整意见由逐份阅读固定输入和给定规则后整理。本脚本只重放这一道题的
参考报告，不是通用合同审查器，也不用于批量制造其他任务。
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


OUTPUT_NAME = "授权委托书风险提示报告.docx"
FONT_SANS = "Noto Sans CJK SC"
FONT_SERIF = "Noto Serif CJK SC"
NAVY = "17365D"
BLUE = "2F5597"
TEAL = "0F6B78"
LIGHT_BLUE = "DCE6F1"
LIGHT_TEAL = "DDEBF7"
LIGHT_GRAY = "F2F2F2"
MID_GRAY = "D9E1F2"
WHITE = "FFFFFF"
RED = "C00000"
AMBER = "9C6500"
GREEN = "548235"


CASES = [
    {
        "number": 1,
        "principal": "张伟",
        "agent": "李芳",
        "case": "房屋租赁合同纠纷｜（2026）京0101民初第0012号",
        "authorization": "一般授权｜2026年4月30日至2027年4月30日",
        "level": "中风险",
        "source_anchors": ["张伟", "李芳", "（2026）京0101民初第0012号"],
        "risks": [
            {
                "code": "M03",
                "name": "受托人无相应资质",
                "evidence": "文件属于诉讼代理；受托人李芳仅列身份证号，未载明律师身份、近亲属关系、单位员工关系或法院许可。",
                "explanation": "诉讼代理人的法定范围受限。现有文本不能证明李芳具备相应诉讼代理资格，代理行为可能不被法院认可。",
                "advice": "补充近亲属、单位员工或法院许可等资格证明；如不符合任一情形，应改由持有效执业证的律师担任受托人。",
            },
            {
                "code": "M07",
                "name": "转委托约定缺失",
                "evidence": "全文未说明是否允许受托人转委托，也未约定转委托条件。",
                "explanation": "受托人能否再次转授权限存在不确定性；未经同意或追认的转委托可能对委托人不发生效力。",
                "advice": "增加“未经委托人书面同意不得转委托”；如允许，应写明范围、次数上限和转受托人资质。",
            },
            {
                "code": "M08",
                "name": "法律责任承担约定缺失",
                "evidence": "正文未约定越权、不当代理或受托人过错造成损失时的责任归属。",
                "explanation": "发生越权或不当代理后，委托人与受托人之间缺少明确的责任分配依据，增加追责难度。",
                "advice": "补充受托人越权行为责任，以及因故意或重大过失造成损失时承担赔偿责任的条款。",
            },
        ],
    },
    {
        "number": 2,
        "principal": "陈雪梅",
        "agent": "王建国律师",
        "case": "商品房买卖合同纠纷｜（2026）浙0101民初第0588号",
        "authorization": "特别授权｜2026年5月1日至2027年2月28日",
        "level": "中风险",
        "source_anchors": ["陈雪梅", "王建国律师", "（2026）浙0101民初第0588号"],
        "risks": [
            {
                "code": "M01",
                "name": "特别授权范围过宽",
                "evidence": "特别授权事项超过3项，并同时包含代为承认和放弃诉讼请求，还包括变更、和解、签收文书和上诉。",
                "explanation": "受托人可以直接处分委托人的实体权利，缺少限制时可能作出对委托人不利且难以撤回的决定。",
                "advice": "逐项保留确有需要的权限；对承认、放弃、变更和和解增加“须经委托人事前书面确认”的限制。",
            },
            {
                "code": "M07",
                "name": "转委托约定缺失",
                "evidence": "全文未说明是否允许转委托，也未约定转委托的范围与条件。",
                "explanation": "转委托权限不清会造成授权链条和行为效力不确定，可能引发连带纠纷。",
                "advice": "明确原则上禁止转委托；确需允许时，限定事项、次数并要求转受托人具备相应律师资质。",
            },
            {
                "code": "M08",
                "name": "法律责任承担约定缺失",
                "evidence": "正文没有约定受托人越权、故意或重大过失时的责任承担。",
                "explanation": "特别授权涉及实体权利处分，缺少责任条款会放大越权或不当代理后的追责风险。",
                "advice": "增加授权边界、重大事项确认流程及故意或重大过失造成损失时的赔偿责任。",
            },
        ],
    },
    {
        "number": 3,
        "principal": "刘洋",
        "agent": "赵明辉",
        "case": "劳动争议｜（2023）粤0305民初第1234号",
        "authorization": "一般授权｜2026年5月1日至2026年5月31日",
        "level": "中风险",
        "source_anchors": ["刘洋", "赵明辉", "（2023）粤0305民初第1234号"],
        "risks": [
            {
                "code": "M03",
                "name": "受托人无相应资质",
                "evidence": "文件属于诉讼代理；赵明辉只列身份证号，未表明律师、近亲属、单位员工或法院许可身份。",
                "explanation": "现有文件不能证明受托人属于可担任诉讼代理人的法定范围，可能影响法院对代理资格的认可。",
                "advice": "补充符合条件的关系或许可证明；若不能证明，应更换为持有效执业证的律师。",
            },
            {
                "code": "M07",
                "name": "转委托约定缺失",
                "evidence": "一般授权条款没有提及是否可以将权限再次转授给第三人。",
                "explanation": "是否允许转委托及其效力边界不清，容易造成未经同意的二次授权争议。",
                "advice": "加入书面同意前置条件；如允许转委托，明确范围、次数和转受托人资质。",
            },
            {
                "code": "M08",
                "name": "法律责任承担约定缺失",
                "evidence": "文件未约定受托人行为后果归属或过错责任。",
                "explanation": "越权或不当代理发生后缺少合同内责任分配依据，委托人追偿成本可能增加。",
                "advice": "明确授权内行为后果、越权责任及受托人故意或重大过失造成损失时的赔偿义务。",
            },
        ],
    },
    {
        "number": 4,
        "principal": "周丽华",
        "agent": "孙志强律师",
        "case": "货物运输合同纠纷｜（2025）京0115民初第0336号",
        "authorization": "一般授权｜2026年4月26日至2026年5月31日",
        "level": "高风险",
        "source_anchors": ["周丽华", "孙志强律师", "（2025）京0115民初第0336号"],
        "risks": [
            {
                "code": "R01",
                "name": "委托人身份证明缺失",
                "evidence": "委托人仅写“周丽华”，未载明身份证号或其他身份证明信息。",
                "explanation": "无法仅凭现有文件核实委托人身份，存在冒名委托、立案材料不完整和代理效力争议。",
                "advice": "补充委托人身份证号，并将身份证明复印件作为附件与本委托书一并留存和提交。",
            },
            {
                "code": "R02",
                "name": "受托人身份证明缺失",
                "evidence": "受托人写为“孙志强律师”，但没有身份证号或律师执业证号。",
                "explanation": "律师身份和诉讼代理资格无法由现有文本核验，法院可能要求补正代理人证件。",
                "advice": "补充律师执业证号、执业机构，并附律师执业证和律所证明等所需材料。",
            },
            {
                "code": "M07",
                "name": "转委托约定缺失",
                "evidence": "正文未约定是否允许孙志强律师转委托。",
                "explanation": "转委托权限和后续授权链条不清，可能导致代理行为效力争议。",
                "advice": "写明未经委托人书面同意不得转委托；如允许，限定范围、次数和资质。",
            },
            {
                "code": "M08",
                "name": "法律责任承担约定缺失",
                "evidence": "正文未约定受托人越权、不当代理或过错造成损失时的责任。",
                "explanation": "缺少责任分配条款会增加越权或不当代理发生后的追责难度。",
                "advice": "增加越权行为责任及故意或重大过失造成损失时的赔偿条款。",
            },
        ],
    },
    {
        "number": 5,
        "principal": "吴强",
        "agent": "黄丽萍",
        "case": "买卖合同纠纷｜（2026）川0107民初第0215号",
        "authorization": "仅写“全权委托”｜2026年4月30日至2026年5月10日",
        "level": "高风险",
        "source_anchors": ["吴强", "黄丽萍", "（2026）川0107民初第0215号"],
        "risks": [
            {
                "code": "R03",
                "name": "授权范围完全未明确",
                "evidence": "正文只写“全权委托”，没有“一般授权”或“特别授权”，也没有逐项列举权限。",
                "explanation": "“全权委托”没有清晰的法律边界，无法判断受托人可实施哪些程序性或实体处分行为。",
                "advice": "改为一般授权或特别授权，并逐项列举允许事项；涉及实体权利处分时增加书面确认限制。",
            },
            {
                "code": "M03",
                "name": "受托人无相应资质",
                "evidence": "文件属于诉讼代理；黄丽萍只列身份证号，未表明律师、近亲属、单位员工或法院许可身份。",
                "explanation": "现有材料不足以证明黄丽萍具备诉讼代理资格，代理行为可能不被法院认可。",
                "advice": "补充符合诉讼代理人范围的资格证明；无法证明时，改由持有效执业证的律师担任。",
            },
            {
                "code": "M07",
                "name": "转委托约定缺失",
                "evidence": "文件没有说明“全权委托”是否包含转委托，也没有任何限制条件。",
                "explanation": "模糊授权叠加转委托空白会扩大权限失控和授权链条不明的风险。",
                "advice": "明确禁止未经书面同意的转委托；如允许，单独列明范围、次数和转受托人资质。",
            },
            {
                "code": "M08",
                "name": "法律责任承担约定缺失",
                "evidence": "正文未写明受托人越权或不当代理的后果与责任。",
                "explanation": "授权边界本就模糊，若再缺少责任条款，发生争议时更难确定双方责任。",
                "advice": "同步补充授权边界、越权责任和故意或重大过失造成损失时的赔偿责任。",
            },
        ],
    },
    {
        "number": 6,
        "principal": "上海盛宏贸易有限公司",
        "agent": "陈明律师",
        "case": "技术服务合同纠纷｜（2026）苏0101民初第0412号",
        "authorization": "特别授权｜2026年6月1日至2026年6月30日",
        "level": "中风险",
        "source_anchors": ["上海盛宏贸易有限公司", "陈明律师", "（2026）苏0101民初第0412号"],
        "risks": [
            {
                "code": "M01",
                "name": "特别授权范围过宽",
                "evidence": "特别授权事项超过3项，并同时包含代为承认和放弃诉讼请求，还含变更、和解与签收法律文书。",
                "explanation": "受托人可以处分公司的重要实体权利；缺少内部确认门槛时，可能产生超出公司真实决策的代理行为。",
                "advice": "收窄到案件必要事项；承认、放弃、变更和和解须经公司有权机构事前书面确认。",
            },
            {
                "code": "M07",
                "name": "转委托约定缺失",
                "evidence": "正文没有说明陈明律师能否转委托，也没有范围和资质限制。",
                "explanation": "公司无法据此控制后续代理人选择与授权链条，可能产生未经同意的二次授权。",
                "advice": "明确禁止未经公司书面批准的转委托；确需允许时，限定事项、次数及律师资质。",
            },
            {
                "code": "M08",
                "name": "法律责任承担约定缺失",
                "evidence": "文件未约定受托人越权、故意或重大过失时的责任。",
                "explanation": "特别授权涉及重大实体权利，缺少责任和追偿条款会增加公司事后救济难度。",
                "advice": "增加越权责任、重大事项审批留痕，以及故意或重大过失造成损失时的赔偿责任。",
            },
        ],
    },
]


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top: int = 80, start: int = 100, bottom: int = 80, end: int = 100) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_width(cell, width_cm: float) -> None:
    width_twips = int(width_cm / 2.54 * 1440)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(width_twips))
    tc_w.set(qn("w:type"), "dxa")


def set_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def set_run_font(run, size: float, *, bold: bool = False, color: str = "000000", font: str = FONT_SANS) -> None:
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    r_fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        r_fonts.set(qn(attribute), font)


def write_cell(cell, text: str, *, size: float = 8.5, bold: bool = False, color: str = "000000", align=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margins(cell)
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.12
    run = paragraph.add_run(text)
    set_run_font(run, size, bold=bold, color=color)


def format_table(table, widths: list[float], *, header: bool = True) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.style = "Table Grid"
    for row_index, row in enumerate(table.rows):
        set_cant_split(row)
        if row_index == 0 and header:
            set_repeat_header(row)
        for column_index, cell in enumerate(row.cells):
            if column_index < len(widths):
                set_cell_width(cell, widths[column_index])


def add_page_field(paragraph) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instruction, separate, placeholder, end):
        run._r.append(element)
    set_run_font(run, 8, color="666666")


def configure_document(document: Document) -> None:
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.55)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.65)
    section.right_margin = Cm(1.65)
    section.header_distance = Cm(0.7)
    section.footer_distance = Cm(0.7)

    normal = document.styles["Normal"]
    normal.font.name = FONT_SANS
    normal.font.size = Pt(10.5)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_SANS)
    normal.paragraph_format.line_spacing = 1.28
    normal.paragraph_format.space_after = Pt(5)

    title = document.styles["Title"]
    title.font.name = FONT_SERIF
    title.font.size = Pt(27)
    title.font.bold = True
    title.font.color.rgb = RGBColor.from_string(NAVY)
    title._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_SERIF)
    title.paragraph_format.space_after = Pt(10)

    for style_name, size, color in (("Heading 1", 17, NAVY), ("Heading 2", 12.5, BLUE)):
        style = document.styles[style_name]
        style.font.name = FONT_SANS
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_SANS)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(8)
        style.paragraph_format.space_after = Pt(6)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = header.add_run("FORTE Legal-020  ·  授权委托书风险提示报告")
    set_run_font(run, 8, color="7F7F7F")

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("第 ")
    set_run_font(run, 8, color="666666")
    add_page_field(footer)
    run = footer.add_run(" 页")
    set_run_font(run, 8, color="666666")


def add_cover(document: Document) -> None:
    spacer = document.add_paragraph()
    spacer.paragraph_format.space_after = Pt(70)

    kicker = document.add_paragraph()
    kicker.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = kicker.add_run("法律文件批量风控核查")
    set_run_font(run, 11, bold=True, color=TEAL)

    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("授权委托书\n风险提示报告")

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(34)
    run = subtitle.add_run("6 份起草稿 · 逐项核查 · 可追溯整改建议")
    set_run_font(run, 12, color="595959")

    meta = document.add_table(rows=4, cols=2)
    format_table(meta, [4.0, 11.5], header=False)
    rows = [
        ("核查依据", "《授权委托书风控校验规则》"),
        ("核查范围", "委托书1 至 委托书6"),
        ("综合分布", "高风险 2 份｜中风险 4 份｜低风险 0 份"),
        ("适用边界", "六份均为起草未签署版本；未签署状态不纳入触发项"),
    ]
    for index, (label, value) in enumerate(rows):
        write_cell(meta.cell(index, 0), label, size=10, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(meta.cell(index, 0), NAVY)
        write_cell(meta.cell(index, 1), value, size=10)
        set_cell_shading(meta.cell(index, 1), LIGHT_GRAY if index % 2 else WHITE)

    note = document.add_paragraph()
    note.paragraph_format.space_before = Pt(30)
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = note.add_run("本报告仅依据所给文件与规则作离线核查，不补造任何外部核验结果。")
    set_run_font(run, 9, color="7F7F7F")
    document.add_page_break()


def add_summary(document: Document) -> None:
    document.add_heading("一、核查结论", level=1)
    paragraph = document.add_paragraph()
    paragraph.add_run("总体判断：").bold = True
    paragraph.add_run("六份文件均存在中风险或高风险项。委托书4因双方身份证明信息缺失、委托书5因授权范围完全未明确，被判定为高风险；其余四份为中风险。")

    table = document.add_table(rows=1, cols=6)
    format_table(table, [1.7, 3.2, 3.0, 2.0, 1.6, 6.0])
    for cell, text in zip(table.rows[0].cells, ["文件", "委托人", "受托人", "等级", "项数", "重点问题"]):
        write_cell(cell, text, size=8.5, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(cell, NAVY)
    for case in CASES:
        row = table.add_row()
        focus = "；".join(risk["name"] for risk in case["risks"])
        values = [f"委托书{case['number']}", case["principal"], case["agent"], case["level"], str(len(case["risks"])), focus]
        for column, (cell, value) in enumerate(zip(row.cells, values)):
            color = RED if case["level"] == "高风险" and column == 3 else AMBER if column == 3 else "000000"
            write_cell(cell, value, size=8.2, bold=column in {0, 3}, color=color, align=WD_ALIGN_PARAGRAPH.CENTER if column in {0, 3, 4} else WD_ALIGN_PARAGRAPH.LEFT)
            set_cell_shading(cell, "FFF2CC" if case["level"] == "中风险" and column == 3 else "FCE4D6" if case["level"] == "高风险" and column == 3 else WHITE)

    document.add_heading("二、核查口径", level=1)
    bullets = [
        "风险等级按已触发项中的最高等级确定：出现任一高风险项即为高风险；没有高风险但出现中风险项即为中风险。",
        "逐份核对主体身份、授权范围、期限、转委托、责任承担和形式要件；只把满足给定触发条件的项目列为风险。",
        "根据原始任务说明，六份均为起草稿，空白签署栏不用于触发本次风险。",
        "未提供的外部证明、律师数据库或法院材料不作推定；调整意见只针对当前文本可以确认的缺口。",
    ]
    for text in bullets:
        paragraph = document.add_paragraph(style="List Bullet")
        paragraph.paragraph_format.space_after = Pt(3)
        paragraph.add_run(text)

    document.add_heading("三、优先整改顺序", level=1)
    priority = document.add_table(rows=4, cols=3)
    format_table(priority, [2.2, 5.5, 9.5])
    for cell, text in zip(priority.rows[0].cells, ["优先级", "对象", "处理动作"]):
        write_cell(cell, text, size=9, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(cell, TEAL)
    priority_rows = [
        ("P0", "委托书4", "先补齐委托人与受托人身份证明，再处理转委托和责任条款。"),
        ("P0", "委托书5", "先把“全权委托”改成明确授权类型和事项，并核实受托人诉讼代理资格。"),
        ("P1", "委托书1、3", "补充受托人资格证明或更换合格代理人。"),
    ]
    for index, values in enumerate(priority_rows, start=1):
        for column, (cell, value) in enumerate(zip(priority.rows[index].cells, values)):
            write_cell(cell, value, size=8.8, bold=column == 0, color=RED if column == 0 else "000000", align=WD_ALIGN_PARAGRAPH.CENTER if column == 0 else WD_ALIGN_PARAGRAPH.LEFT)
            set_cell_shading(cell, LIGHT_GRAY if index % 2 == 0 else WHITE)


def add_case_section(document: Document, case: dict) -> None:
    heading = document.add_heading(f"委托书{case['number']}｜{case['principal']} → {case['agent']}", level=1)
    # 让案例紧接上一页剩余空间自然排版；summary 的优先整改表可能跨页，
    # 强制段前分页会把最后一行表格单独留在一页并制造近空白页。
    heading.paragraph_format.page_break_before = False
    heading.paragraph_format.keep_with_next = True

    facts = document.add_table(rows=3, cols=2)
    format_table(facts, [3.1, 14.0], header=False)
    fact_rows = [
        ("案件", case["case"]),
        ("授权", case["authorization"]),
        ("来源文件", f"委托书{case['number']}.docx"),
    ]
    for index, (label, value) in enumerate(fact_rows):
        write_cell(facts.cell(index, 0), label, size=9, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(facts.cell(index, 0), BLUE)
        write_cell(facts.cell(index, 1), value, size=9.2)
        set_cell_shading(facts.cell(index, 1), LIGHT_GRAY if index % 2 else WHITE)

    conclusion = document.add_table(rows=1, cols=2)
    conclusion.alignment = WD_TABLE_ALIGNMENT.CENTER
    conclusion.autofit = False
    set_cell_width(conclusion.cell(0, 0), 3.1)
    set_cell_width(conclusion.cell(0, 1), 14.0)
    write_cell(conclusion.cell(0, 0), "风险等级", size=10, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_shading(conclusion.cell(0, 0), NAVY)
    level_color = RED if case["level"] == "高风险" else AMBER
    write_cell(conclusion.cell(0, 1), f"{case['level']}｜共 {len(case['risks'])} 项", size=11, bold=True, color=level_color, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_shading(conclusion.cell(0, 1), "FCE4D6" if case["level"] == "高风险" else "FFF2CC")

    document.add_heading("风险明细与调整意见", level=2)
    risk_table = document.add_table(rows=1, cols=4)
    format_table(risk_table, [3.2, 4.3, 4.7, 5.1])
    for cell, text in zip(risk_table.rows[0].cells, ["风险项", "触发证据", "风险说明", "调整意见"]):
        write_cell(cell, text, size=8.5, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(cell, NAVY)
    for risk in case["risks"]:
        row = risk_table.add_row()
        values = [f"{risk['code']}｜{risk['name']}", risk["evidence"], risk["explanation"], risk["advice"]]
        for column, (cell, value) in enumerate(zip(row.cells, values)):
            write_cell(cell, value, size=8.1, bold=column == 0, color=RED if risk["code"].startswith("R") and column == 0 else AMBER if column == 0 else "000000")
            set_cell_shading(cell, "FFF2CC" if column == 0 and risk["code"].startswith("M") else "FCE4D6" if column == 0 else WHITE)

    trace = document.add_paragraph()
    trace.paragraph_format.space_before = Pt(6)
    run = trace.add_run("核查边界｜")
    set_run_font(run, 8.5, bold=True, color=TEAL)
    run = trace.add_run("本节只列给定规则下实际触发的项目；该文件为起草稿，空白签署栏未纳入风险触发。")
    set_run_font(run, 8.5, color="595959")


def validate_fixed_inputs(input_dir: Path) -> None:
    rules = (input_dir / "授权委托书风控校验规则.md").read_text(encoding="utf-8")
    for case in CASES:
        path = input_dir / "委托书" / f"委托书{case['number']}.docx"
        document = Document(path)
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        missing = [anchor for anchor in case["source_anchors"] if anchor not in text]
        if missing:
            raise ValueError(f"{path.name} 与 Legal-020 固定输入不符，缺少 {missing}")
        for risk in case["risks"]:
            if risk["name"] not in rules:
                raise ValueError(f"规则文件缺少风险项：{risk['name']}")


def solve(input_dir: Path, output_path: Path) -> None:
    validate_fixed_inputs(input_dir)
    document = Document()
    document.core_properties.title = "授权委托书风险提示报告"
    document.core_properties.subject = "FORTE Legal-020 六份授权委托书逐项风控核查"
    configure_document(document)
    add_cover(document)
    add_summary(document)
    for case in CASES:
        add_case_section(document, case)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir / OUTPUT_NAME)


if __name__ == "__main__":
    main()
