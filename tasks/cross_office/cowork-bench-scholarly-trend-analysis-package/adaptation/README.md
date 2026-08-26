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
