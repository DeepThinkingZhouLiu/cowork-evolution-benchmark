请基于以下 5 个冻结输入，制作一份 6 页、面向 `strategy@company.example.com` 的 StrataCorp v4 competitive response brief：

- `/app/input_files/Competitive_Intelligence_Tracker.xlsx`
- `/app/input_files/Analyst_Note_StrataCorp_v4.eml`
- `/app/input_files/Pricing_Page_Change_Policy.eml`
- `/app/input_files/Competitive_Intel_Slack.json`
- `/app/input_files/AutomationBench_1115_Context.json`

必须满足：

- 把追踪表的 6 个 feature 做成原生可编辑矩阵，完整保留 StrataCorp、我方、gap status、priority 和 notes 映射。
- 用原生可编辑图表汇总 gap status：Parity=2、Behind=2、Ahead=1、Gap=1。
- 分源呈现分析师邮件的关键事实：AI 报告生成 `3x` 更快、移动端 `4.2/5`、Starter `$29/user/mo`、Pro `$79/user/mo`、Enterprise `custom pricing`、价格低 `15-20%`、工作流只支持 `12 trigger types` 对比行业 `30+`、Embedded Analytics 已 GA。
- 分源呈现 Slack 现场信息：
  - Meridian Corp：Pro tier 切换报价折扣 `30%`，deal size `~$45K ARR`。
  - 两个 prospects 提到新 mobile app，其中一个称它是 deciding factor。
  - NovaTech：喜欢 AI insights，但认为 custom workflows `basically nonexistent`；如我方发布 AI module，可能回流。
- 明确形成响应优先级：对 AI-Powered Insights 的 P0 差距、Native Mobile App 的 P1 gap、Embedded Analytics 的 P1 behind 做应对；同时用 Custom Workflows 的 Ahead 优势制作 battle card。
- 必须有一页价格治理/执行计划：原样保留三个价格 tier，标明不得直接修改 pricing comparison page；应创建 subject 为 `Pricing Page Update Request` 的 HubSpot ticket，交由 Pricing Committee 周三审议。
- 不得把价格写入竞争追踪表 feature 行，不得混淆表格、邮件和 Slack 来源，不得编造其他市场数字。
- 文稿保持管理咨询风格，所有文本、表格和图表原生可编辑，无重叠、无裁切、无整页压平。

输出文件必须命名为 `/app/output/StrataCorp_v4_Competitive_Response_Brief.pptx`。请仅在 `/app/output` 中写入最终交付文件。
