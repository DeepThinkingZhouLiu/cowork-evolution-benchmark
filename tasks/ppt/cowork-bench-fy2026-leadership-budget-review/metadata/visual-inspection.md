# 最终 Golden 目视检查记录

- 检查时间：2026-08-24（Asia/Shanghai）
- 被检 PPTX：`solution/budget_vs_actuals.pptx`
- PPTX SHA256：`72134e27ca0db73f15e79af30d566cbdfab107cae8ebd6a42b6604e3fda95e71`
- 渲染引擎：LibreOffice 7.3.7.2 30(Build:2)
- 被检 PDF：`solution/qa/budget_vs_actuals-rendered.pdf`
- PDF SHA256：`757bd907e14d33ea42a69a7ce62e0ca971f00bf645b8e5ae4470471828f7c4e2`
- PDF 页面：10 页，959.981 × 540 pt，16:9
- 检查方式：先检查完整 contact sheet，再以 150 DPI 逐页打开 `solution/qa/pdf-pages/page-01.png` 至 `page-10.png`；最终重建后的 10 张页面 PNG 与已逐页查看的页面逐字节一致，未只依赖程序化重叠检查。

## 逐页结果

- 第 1 页：标题与副标题完整；深色底上的文字和品牌图形清晰；无截断、遮挡或越界。
- 第 2 页：7 个部门、预算总额、实际总额、差异总额均清晰；表格行列完整；来源脚注未截断。
- 第 3 页（Инженерия）：部门名、状态、预算、人数、平均薪资、实际支出、差异和百分比均清晰；左右栏无冲突。
- 第 4 页（Кадры）：全部指标清晰；左右栏、分隔线和底部状态条无重叠。
- 第 5 页（НИОКР）：全部指标清晰；俄文标题正常显示；无字体替换导致的溢出。
- 第 6 页（Операции）：全部指标清晰；金额和两位小数完整显示。
- 第 7 页（Поддержка）：全部指标清晰；最长部门名仍有安全边距。
- 第 8 页（Продажи）：全部指标清晰；来源脚注与状态条均在页面内。
- 第 9 页（Финансы）：全部指标清晰；标题、正文和状态条无重叠。
- 第 10 页：Most Under Budget 与 Most Over Budget 两组结论、对应数值以及“全部 7 个部门均低于预算”的总结均清晰，左右栏一致。

## 结论

10 页全部通过。没有发现元素重叠、裁切、缺字、乱码、不可读的小字、异常换行或页面尺寸错误。PPTX 内容采用原生可编辑形状和文本，不是整页图片。

## XLSX 打印版检查

- 被检 XLSX：`solution/budget_vs_actuals.xlsx`
- XLSX SHA256：`744cdf99e9b343d19d9c97043478554e31eb606a4e7cc127ffd209c9d9145d97`
- 被检 PDF：`solution/qa/budget_vs_actuals-xlsx-rendered.pdf`
- PDF SHA256：`042da8b0691d2cf5c64f4b8cc1f0a6ad472a0d21cfb67e832b4a3afd65c16dc8`
- PDF 页面：2 页，A4；`Department Details` 为横向单页，`Totals` 为纵向单页。
- 第 1 页：7 列和 7 个部门均完整出现在同一页；俄文部门名、金额、百分比、人数和平均薪资清晰，无横向拆页、截断或遮挡。
- 第 2 页：6 项 Totals 指标和值完整、清晰；金额与俄文部门名正常显示。
- 结论：XLSX 两张工作表的打印版全部通过目视检查。
