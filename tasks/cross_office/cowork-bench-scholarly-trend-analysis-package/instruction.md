请使用工作区中的上游学术数据快照和 HTML 数据面板，完成“抽取 -> 关联 -> 分析 -> 呈现”任务，并交付：

- `/app/output/scholarly_trend_processor.py`
- `/app/output/scholarly_trend_results.json`
- `/app/output/Trend_Analysis_Report.xlsx`
- `/app/output/Trend_Analysis_Presentation.pptx`

这是 Cowork-Bench `pw-scholarly-trend-analysis-excel-ppt` 的离线静态改造版。上游原题要求访问 scholarly 服务和 `localhost:30324`；本题把同一上游数据固化为本地 JSON 和 HTML，不需要网络或外部服务。

输入责任：

- `scholarly_papers.json`：上游 preprocess 注入 scholarly 数据库的 4 篇论文快照，包含 `Paper_ID, Title, Area`。
- `dashboard.html`：上游 `localhost:30324` 数据面板的原始 HTML，表格包含 `Research_Area, Publication_Count_2025, Avg_Citations, Trending_Score`。
- `analysis_guide.md`：上游主任务的匹配口径，其中 `Citations = Avg_Citations`，`Relevance_Score = Trending_Score`。
- `Analysis_Guidelines.pdf`、`Report_Template.xlsx`：上游附带的参考材料。PDF 中还提到 gap 分析，但本题没有提供可与 benchmark 对比的内部指标，因此不得编造 gap 或趋势方向。

`scholarly_trend_processor.py` 必须真实读取 JSON 和 HTML、按 `Area = Research_Area` 精确关联、按 Area 字母序排序，并支持以下命令行：

```bash
python scholarly_trend_processor.py \
  --papers /app/data/input_files/scholarly_papers.json \
  --dashboard /app/data/input_files/dashboard.html \
  --output /app/output/scholarly_trend_results.json
```

JSON 必须包含以下结构：

- `data_analysis`：4 行，每行恰好包含 `Paper_ID, Title, Area, Citations, Relevance_Score`。
- `metrics`：包含 `Total_Papers, Avg_Citations, Top_Area, Avg_Relevance`。两个平均数按标准四舍五入取整。
- `recommendations`：至少 2 行，每行包含 `Priority, Action, Area`，Area 必须来自 `data_analysis`。

Judge 会用原始输入和一份修改后的 HTML 分别执行该脚本，所以不得把面板数字直接硬编码成固定 JSON。

`Trend_Analysis_Report.xlsx` 必须包含三张可编辑工作表：

- `Data_Analysis`：表头恰为 `Paper_ID, Title, Area, Citations, Relevance_Score`，4 行来源匹配结果，按 Area 字母序排列。
- `Metrics`：表头为 `Metric, Value`，依次包含 `Total_Papers, Avg_Citations, Top_Area, Avg_Relevance`。四个 Value 必须用 Excel 公式从 `Data_Analysis` 推导。
- `Recommendations`：表头为 `Priority, Action, Area`，与 JSON 的建议一致，至少 2 行。

`Trend_Analysis_Presentation.pptx` 使用 16:9，至少 4 页，用俄文讲清数据来源、匹配方法、4 篇论文的结果、四项汇总指标和建议。所有关键文字、数字、图表和形状必须可编辑，不得以整页截图充当幻灯片。Excel、JSON 和 PPT 的论文、领域和指标必须一致。
