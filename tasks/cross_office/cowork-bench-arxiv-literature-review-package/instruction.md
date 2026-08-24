请基于工作区中的 Cowork-Bench 上游数据快照，完成一次可追溯的 Arxiv 文献筛选、外部趋势匹配和多文件交付。请将以下 5 个交付物保存到指定路径：

- `/app/output/arxiv_litreview_processor.py`
- `/app/output/arxiv_litreview_results.json`
- `/app/output/Lit_Review_Report.xlsx`
- `/app/output/Lit_Review_Analysis.docx`
- `/app/output/Arxiv_Litreview_Tracker.xlsx`

这是 Cowork-Bench `fetch-arxiv-lit-review-word-gsheet` 的逐题离线改造版。上游原题依赖本地 Arxiv/PostgreSQL、`localhost:30323` 和 Google Sheet 服务；本题将同一批上游记录静态化为 JSON，并把 Google Sheet 交付替换为离线可编辑的 Tracker Excel，不需要网络或外部服务。

输入责任：

- `arxiv_papers.json`：上游 `preprocess/main.py` 向 `arxiv.papers` 注入的 5 条记录的完整字段投影。其中 4 条主分类以 `cs.` 开头，1 条为 `quant-ph` 噪声。
- `benchmark_data.json`：上游 `localhost:30323/api/data.json` 的原始响应，包含 `LLMs / Computer Vision / RL` 的外部增长率。
- `analysis_guide.md`、`Analysis_Guidelines.pdf`、`Report_Template.xlsx`：上游原始参考附件。指南要求 gap 和季度趋势，但实际输入没有内部基准、引用数、相关性分数或时间序列，所以不得编造这些指标。

筛选与匹配口径：

- 仅保留 `primary_category` 以 `cs.` 开头的记录；其他记录放入 JSON 的 `excluded_papers`，不得进入报告主表和 Tracker。
- 将 `title + summary` 转为小写文本后按以下透明规则匹配：包含 `language model` -> `LLMs`；包含 `vision transformer` -> `Computer Vision`；包含 `reinforcement learning` -> `RL`；其余为 `Unmatched`。
- `Growth_Pct` 只能来自 `benchmark_data.json` 中同名领域，不匹配时留空。外部增长率不等于论文本身的增长、质量或引用表现。
- 主表按 `Paper_ID` 升序排列。平均值保留两位小数，覆盖率按百分比计算。

`arxiv_litreview_processor.py` 必须真实读取两份 JSON 并支持：

```bash
python arxiv_litreview_processor.py \
  --papers /app/data/input_files/arxiv_papers.json \
  --benchmark /app/data/input_files/benchmark_data.json \
  --output /app/output/arxiv_litreview_results.json
```

`arxiv_litreview_results.json` 必须包含：

- `source_summary`：输入论文数、保留数、排除数和外部主题数。
- `data_analysis`：4 条保留记录，字段为 `Paper_ID, Title, Primary_Category, Published, Benchmark_Area, Growth_Pct, Match_Status`。
- `excluded_papers`：被排除记录的 `Paper_ID, Title, Primary_Category, Exclusion_Reason`。
- `metrics`：`Total_Papers, Excluded_Noise_Papers, Benchmark_Matched_Papers, Avg_Matched_Growth_Pct, Top_Growth_Area, Benchmark_Coverage_Pct`。
- `recommendations`：对已匹配的 3 个领域按 `Growth_Pct` 降序排列，包含 `Rank, Priority, Action, Area, Growth_Pct`。
- `limitations`：说明缺失的引用数、相关性、内部基准和时间序列边界。

Judge 会在原始输入与修改后的输入上分别执行处理脚本，所以不得将 Paper_ID、增长率或汇总结果直接写死在结果文件中。

`Lit_Review_Report.xlsx` 必须包含 3 张可编辑工作表：

- `Data_Analysis`：表头与 JSON `data_analysis` 字段一致，包含 4 行，不得出现噪声记录。
- `Metrics`：表头为 `Metric, Value`，包含上述 6 项指标。除排除数外，可从 `Data_Analysis` 推导的 5 项必须使用 Excel 公式；排除数来自输入筛选结果。
- `Recommendations`：表头为 `Rank, Priority, Action, Area, Growth_Pct`，与 JSON 的 3 条建议一致。

`Arxiv_Litreview_Tracker.xlsx` 是上游 Google Sheet 的离线替代，工作表名为 `Arxiv Litreview Tracker`，包含 4 条保留论文的 `Paper_ID, Title, Primary_Category, Benchmark_Area, Growth_Pct, Review_Status`，并与 JSON/Report 一致。

`Lit_Review_Analysis.docx` 正文使用俄文，至少包含 `Краткое резюме`、`Источники и метод`、`Ключевые выводы`、`Рекомендации`、`Ограничения` 等结构，讲清 5 -> 4 -> 3 的筛选与匹配链路、4 篇保留记录、3 个外部增长率、建议与数据边界。不得将上游 Golden 中无法从输入推导的引用数和相关性分数当作事实。
