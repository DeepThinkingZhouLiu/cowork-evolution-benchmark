# 逐题改造记录

本题只改造 Cowork-Bench `pw-scholarly-trend-analysis-excel-ppt`，不是按 `extract_analyze_present` 名称自编的任务。

- 上游仓库：`0717376/cowork_bench`
- 上游 commit：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 上游任务路径：`tasks/finalpool/pw-scholarly-trend-analysis-excel-ppt`
- 许可：Apache-2.0

原题的工作链是：从 scholarly 数据库取4篇论文 -> 从本地网页抽取5个领域指标 -> 按领域关联 -> 写处理脚本和 JSON -> 交付三表 Excel 和 PPT。本地只将数据库和 HTTP 服务固化为 JSON/HTML，任务意图、匹配规则、文件名、列名和页面要求都来自上游。

逐项阅读后确认的来源真值：

- Computer Vision / `Vision Transformer Advances`：Citations `10.8`，Relevance `82`
- Federated Learning / `Federated Learning Survey`：Citations `7.2`，Relevance `68`
- Large Language Models / `Scaling Laws for LLMs`：Citations `12.3`，Relevance `95`
- Reinforcement Learning / `Multi-Agent RL Framework`：Citations `8.5`，Relevance `71`
- 汇总：Total Papers `4`，Avg Citations 取整为 `10`，Top Area 为 `Large Language Models`，Avg Relevance 为 `79`

上游 Evaluator 只检查 Agent 自身 Excel 内部一致，没有把面板数字与输入真值核对，因而一份“整体编错但自洽”的交付也可能通过。本地 Judge 直接解析 JSON/HTML 重算，还会改动 HTML 的 Computer Vision 指标再回放交付脚本，用于拒绝硬编码。

`Analysis_Guidelines.pdf` 的 gap 分析建议与本题可用数据不完整对齐；因没有内部对比值，本地不要求也不允许编造 gap。

## v2 连续评分与视觉校准

- `rubric.v1.json` 与 `tests/judge_v1.py` 原样保留第一版严格通过/失败口径；`SCORING_VERSION=v1|v2` 可显式复现，默认使用 v2。
- v2 只有脚本、JSON、Excel 或 PPT 缺失、为空、损坏或不可解析时硬归零；其余按处理脚本原始/变异回放、4×5 个论文字段、4 项指标、2 条建议、Excel 公式、PPT 内容和跨文件一致性逐原子单元给连续部分分。
- 任务专属负例共 38 个：14 个缺失/空/损坏/原样输入致命反例均为 0；24 个单事实、公式、版式、跨文件冲突和严重但可解析的不完整反例落在 `0.587755–0.997386`，避免把“只错一个字段”和“几乎没做”都压成 0。
- 本题含 `Trend_Analysis_Presentation.pptx`，因此额外启用 `visual_rubric.json`。VLM 只评价五页研究简报的叙事、论文卡片/排名比较、三步关联流程、首要发现和全稿视觉一致性；开放式生成采用 `reference_free`，不要求复制 Golden 版式。
- 固定 Codex CLI `0.148.0` 与 the configured VLM provider 的 5 次 Golden 视觉分为 `0.978, 1.000, 1.000, 1.000, 1.000`，均值 `0.9956`、总体标准差 `0.0088`。任务专属中等负例仅用遮罩盖住第3页一张论文卡片，文件仍可打开且代码分仍为 `1.0`，视觉分降为 `0.857`，按 70/30 聚合后为 `0.9571`。
- 视觉/API 调用最终基础设施失败为 0；后续若渲染或模型调用失败，必须记录为 `infrastructure_failure`，不得按 0 分或静默退化为纯代码分。
