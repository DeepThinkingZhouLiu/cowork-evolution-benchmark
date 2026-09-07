# CS161 名册控制手册

你是系办的课程数据管理员。请以 `/app/input_files/CS161_Class_Roster.xlsx` 为唯一事实来源，制作一份供教务组复核名册总数、源顺序和评分字段的 Word 控制手册。

只向 `/app/output` 交付一个可编辑文件：`cs161_roster_control_book.docx`。不要交付 PDF、截图、XLSX 副本或中间文件。

请完成以下内容：

- 制作封面，标题为 `CS161 Roster Control Book`，标明来源文件、工作表 `Sheet1` 和数据行 `2–10`；紧接一个可更新的 Word 目录字段，覆盖文档中的一、二级标题。
- 在 `Control Summary` 中写明源数据行数，并给出专业和年级分布。分布必须由输入现场计数，并同时列出人数与占总人数的百分比（保留 1 位小数）。
- 在一个单独的横向页节中制作 `Source-order Roster Register`。使用原生 Word 表格，按源表顺序精确转录 9 名学生的 `Source row / Student ID / Last name / First name / Academic standing / Major / HW1 / HW2 / Midterm / HW3 / Final`。`Source row` 使用 `Sheet1!2`至 `Sheet1!10`；每个 Student ID 所在行设置独立书签，供后文跳转。
- 在后续纵向页节的 `Assessment Reconciliation` 中，对 `HW1 / HW2 / Midterm / HW3 / Final` 分别列出有效人数、平均分（保留 2 位小数）、最低分和最高分，全部从输入重算。
- 在 `Data-quality Controls` 中检查并用原生表格记录：9 个数据行是否全部填满、Student ID 是否唯一、是否全部符合 `S` 加 4 位数字、45 个评分是否全部为 0–100 的整数。每项写明 `PASS` 或 `FAIL` 及可复核的计数证据。
- 在 `Alphabetical Name Index` 中按 Last name、First name 升序列出全部 9 人，每行给出 `Name / Student ID / Source row / Roster link`；`Roster link` 必须是跳回横向名册对应学生行的内部链接，不是网址或普通文本。
- 在文档末尾给出对账句：`Control total: 9 source rows = 9 roster rows = 9 index entries.`，其中数字必须按输入与交付内容现场确定。
- 名册、分布、评分对账、数据质量和索引均使用可编辑的原生 Word 表格；表头需重复，数据行不得拆分跨页。三个页节使用一致的文档标题页眉和连续页码。不得使用宏、OLE、外部链接或整页图片代替可编辑内容。

这份手册的任何名单、计数、比例和统计都必须能回到冻结的 Excel 单元格核对；不要推测学生背景或补写输入中不存在的政策。
