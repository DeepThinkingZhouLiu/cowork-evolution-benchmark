请基于下列冻结输入，完成股价数据清洗、董事会演示文稿和演讲稿：

- `/app/input_files/Original_Request.md`
- `/app/input_files/stock_data_raw.csv`

仅在 `/app/output` 中交付：

- `stock_cleaned.csv`
- `stock_analysis.pptx`
- `stock_analysis_script.md`

数据清洗必须可确定复现：

- 原始文件共 3,740 行。8 列，保留原列名和列顺序。
- 删除 5 行带 `FakeDate_0` 至 `FakeDate_4` 且必需字段缺失的无效记录，不得填补或伪造这些值。
- 按 `Date` 升序排列；结果恰好 3,735 行，不得有空值。
- `Unnamed: 0` 和 `Volume` 保存为整数，其余数值列保留原始精度。

`stock_analysis.pptx` 必须是恰好 5 页、16:9，一页一个原生可编辑的数据可视化，页序和标题为：

- `Closing Prices Over Time`：折线图，展示 Close 的长期趋势。
- `Close Price with 50 and 200-Day Moving Averages`：同图展示 Close、50-Day MA 和 200-Day MA。
- `Correlation Between Columns`：可编辑热力表，覆盖 `Adj_Close` / `Close` / `High` / `Low` / `Open` / `Volume` / `Daily_Return` 的完整 7×7 相关矩阵。
- `Daily Return Distribution`：可编辑柱状直方图。
- `Volume Distribution`：可编辑柱状直方图。

计算和显示口径：

- 所有统计量均基于全部 3,735 条日度记录；`Daily_Return = Close.pct_change()`，年化波动率使用日收益率样本标准差乘以 `sqrt(252)`。
- 两张时间序列图为避免 3,735 个点在页面上过密，使用每个自然月的最后一条记录显示，共 179 个月度点；MA 先按日度数据计算，再取月末值。
- 两张分布图分别对 3,734 个非空日收益率和 3,735 个 Volume 使用 NumPy 等宽 30 bins 直方图。

每页必须包含简短董事会解读、关键数字和来源。不得伪造股票代码、事件原因、预测、买卖建议或因果结论；应显式说明历史趋势不等于预测、相关不等于因果。图表与热力表要原生可编辑，不得以截图或整页图片代替。

`stock_analysis_script.md` 必须向 board members 问候并交代会议背景，以 `Slide 1` 至 `Slide 5` 与演示文稿一一对应，准确解释每图的趋势、业务含义和证据边界，并以感谢听众出席会议的结语收尾。

成品必须能在 LibreOffice Impress 中打开和渲染，无重叠、裁切、超出页面、宏或外部文件依赖。
