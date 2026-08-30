请阅读 `/app/input_files/financial_report.pdf`，制作一份 6 页、16:9 的 NVIDIA Q1 FY2024 管理层财务简报。

必须满足以下要求：

- 只使用输入 PDF 中的事实，不补充外部股价、估值、预测或投资建议。
- 用原生可编辑图表展示 Q1 FY23、Q4 FY23、Q1 FY24 三期 Revenue，数值分别为 8,288、6,051、7,192 百万美元。
- 用原生可编辑表格完整展示 GAAP Q1 FY2024 Summary 的 Revenue、Gross margin、Operating expenses、Operating income、Net income、Diluted EPS，以及三期数值和 Q/Q、Y/Y 变化。
- 分业务展示 Data Center、Gaming、Professional Visualization、Automotive 的本季度收入和 Q/Q、Y/Y 变化；必须明确 Data Center 和 Automotive 为 record quarter。
- 明确区分 GAAP 与 Non-GAAP，准确展示本季度 Gross margin、Operating expenses、Operating income、Net income 和 Diluted EPS；不得把 Non-GAAP 数值冒充 GAAP。
- 展示 Operating cash flow 2,911 百万美元、Free cash flow 2,643 百万美元和 Cash dividends 99 百万美元。
- 完整展示 Q2 FY2024 outlook：Revenue 11.00 billion ±2%；GAAP / Non-GAAP gross margin 68.6% / 70.0%，均 ±50 bps；GAAP / Non-GAAP operating expenses 2.71 / 1.90 billion；other income 90 million；tax rate 14.0% ±1%。
- 对 outlook 明确标注 forward-looking，并说明不是未来表现保证；对 Non-GAAP 明确说明其不能替代 GAAP。
- 每页标注输入文件和对应页码来源。推导性结论必须标记为 derived reading，且不得超出输入支持范围。
- 所有文本、表格和图表必须原生可编辑，不得把页面压成整页图片；版面无重叠、无裁切，适合管理层阅读。

输出文件必须命名为 `/app/output/NVIDIA_Q1_FY2024_Management_Briefing.pptx`。请仅在 `/app/output` 中写入最终交付文件。
