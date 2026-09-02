你是学籍数据审计专员。请读取 `/app/input_files/transcripts.pdf`，把可见的学业记录逐项转换为可复算的 Excel 审计工作簿，并交付一份与工作簿一致的 Word 审计报告。PDF 为图像型文件；只允许使用实际交付的这份 PDF，不得访问网络、猜测未交付页内容，也不得自行套用未提供的加权 GPA 政策。

在 `/app/output` 中交付且只交付以下两个文件：

1. `Transcript_GPA_Audit.xlsx`，依次包含四个工作表：
   - `Source_Profile`：记录学生姓名、签发日期、表头中的 Credits Earned、Weighted GPA 和 Unweighted GPA，同时记录实际 PDF 页数、页脚印刷标记和来源完整性风险。不要把地址、电话或出生信息复制到交付物中。
   - `Course_Ledger`：按 PDF 中的学年、学期和阅读顺序列出全部可见课程，字段为 Sequence、Academic_Year、Term、Course、Program_Tag、Grade、Credits、Unweighted_Grade_Points 和 Unweighted_Quality_Points。`Program_Tag` 只保留课程名中明示的 `DE` 或 `AP`；后两个计算字段必须使用可复制的原生公式，未加权换算固定为 A=4、B=3、C=2、D=1、F=0。
   - `Term_Audit`：按每个可见学期汇总 Course_Count、Credits、Unweighted_Quality_Points、Calculated_Unweighted_Term_GPA、Displayed_Term_GPA、Cumulative_Credits、Cumulative_Unweighted_Points、Calculated_Cumulative_Unweighted_GPA、Displayed_Cumulative_GPA 和 Displayed_vs_Calculated_Delta。除 PDF 中直接读取的两个 `Displayed` 字段外，其他计算字段必须使用原生公式。
   - `Summary`：用公式给出课程数、学分数、A/B/C 数量、未加权质量分、重算未加权 GPA、与表头报告值的差异、最终显示累计 GPA 与表头加权 GPA 的差异，并用公式输出来源完整性标记。提供一个可编辑折线图，对比按学期重算的累计未加权 GPA 与 PDF 显示的累计 GPA。
2. `Transcript_GPA_Audit_Report.docx`，标题为 `Transcript GPA Audit Report`，至少包含 `Source Completeness`、`Reported Academic Profile`、`Course Extraction and Recalculation`、`Term-Level Reconciliation`、`Weighted GPA Evidence Boundary` 和 `Audit Conclusion` 六节。报告必须：
   - 明确区分“PDF 实际交付页数”和“页脚印刷的页码声明”，并把两者不一致写成来源完整性风险。
   - 准确报告课程数、学分、成绩分布、未加权质量分、重算 GPA 和表头 GPA。
   - 用原生表格列出每个学期的重算结果，并单独标明最后一个学期的显示值与未加权重算值差异。
   - 把 PDF 中报告的加权 GPA 作为“来源报告值”保留，但明确说明缺少加权政策和未交付页，因此不能独立重建或认定其计算规则。

Excel 与 Word 中的人名、页数、课程、成绩、学分、GPA、差异和风险结论必须完全一致；不得用截图替代可编辑的表格或图表。
