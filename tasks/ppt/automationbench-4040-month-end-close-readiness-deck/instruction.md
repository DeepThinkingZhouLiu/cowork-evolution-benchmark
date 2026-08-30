你是财务运营团队的演示文稿分析师。请基于以下冻结输入制作一份 5 页、面向 Controller 和 CFO 的 January 2026 month-end close readiness 简报：

- `/app/input_files/January_2026_Close_Checklist.xlsx`
- `/app/input_files/AutomationBench_4040_Context.json`

必须满足：

- 仅使用冻结输入中的事实，不得编造金额、任务、负责人、状态或截止日期。
- 文稿必须清楚表达总体状态 `Not Ready`：7 项中 4 项 `Complete`，3 项未完成，完成率为 57.1%。
- 必须给出一张原生可编辑的完成/未完成图表，数值为 4 和 3，不得用截图代替。
- 必须给出一张原生可编辑明细表，完整包含 7 个 checklist item 的 Item、Owner、Status 和 Due，不得遗漏或错配。
- 必须单独标出 3 个未完成项：
  - `Revenue accruals` — `Carol Diaz` — `In Progress` — `2026-02-02`。
  - `Intercompany elimination` — `Dave Kim` — `Not Started` — `2026-02-02`。
  - `Tax provision estimate` — `Grace Wu` — `Blocked` — `2026-02-03`，并保留备注 `Waiting on state tax guidance`。
- 冻结时间为 `2026-02-03T16:00:00Z`，因此前两个未完成项应标为 overdue，Tax provision 应标为 due today / blocked，不得改变原 Status 字段。
- 末页给出可供沟通使用的准备绪论，并明确标注受众 `controller@company.example.com`、`cfo@company.example.com` 以及 Slack 频道 `finance-alerts`。
- 使用一致的财务管理风格，保持文本、图表和表格原生可编辑；页面无重叠、无裁切、无整页压平。

输出文件必须命名为 `/app/output/January_2026_Close_Readiness_Briefing.pptx`。请仅在 `/app/output` 中写入最终交付文件。
