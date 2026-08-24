# 逐题改造记录

本题只改造 Cowork-Bench `fetch-arxiv-lit-review-word-gsheet`，没有从分类名或模板自行生成数据、论文或 Query 故事。

- 上游仓库：`0717376/cowork_bench`
- 上游 commit：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 上游路径：`tasks/finalpool/fetch-arxiv-lit-review-word-gsheet`
- 许可：Apache-2.0
- 原始任务：从本地 Arxiv 找出 4 篇 AI/ML 记录，排除 1 篇 `quant-ph` 噪声，合并 localhost 趋势数据，交付脚本、JSON、Excel、Word 和 Google Sheet。

人工阅读的上游文件包括：

- `docs/task.md` 和 `docs/agent_system_prompt.md`
- `preprocess/main.py` 中的 5 条 `arxiv.papers` 与 2 条 `arxiv_latex.papers` 注入
- `evaluation/main.py`
- `files/mock_pages.tar.gz` 与其中的 `api/data.json`
- `initial_workspace` 的 Markdown、PDF 和 Excel
- `groundtruth_workspace` 的 Word 与 Excel
- 仓库根目录 `LICENSE`

静态化只做了两件事：

- 把 `preprocess/main.py` 中 `arxiv.papers` 的 5 条记录按原值写入 `arxiv_papers.json`。
- 把上游 API 响应原文件复制为 `benchmark_data.json`，并用离线 `Arxiv_Litreview_Tracker.xlsx` 替代不可回放的 Google Sheet 服务。

可直接从输入复算的真值是：

- 论文记录总数 `5`，保留 `4`，排除 `1`。
- 按透明关键词规则匹配的领域为 `LLMs=45%`、`Computer Vision=12%`、`RL=8%`；Federated Learning 没有同名 benchmark 主题。
- 匹配论文数 `3`，平均外部增长率 `21.67%`，匹配覆盖率 `75%`，最高增长领域 `LLMs`。

上游 Golden 自行加入了引用数 `450/320/180/250`、平均引用数 `300` 和相关性分数 `95/82/71/68`；这些数字在 preprocess、API 响应和附件中都不存在。它们作为上游缺陷证据保留，不进入本地 Golden，Judge 也会拒绝这类无来源断言。

需要明确的边界：上游论文记录本身是 benchmark 注入数据，不应对外宣称为经网络验证的真实 Arxiv 论文。本题的“真实来源”指完整取自已授权的上游 benchmark 任务，不是本项目自编输入。

验证记录：

- Golden 完整 Judge 得分 `1.0`；处理脚本在 Paper_ID 和两个 Growth_Pct 被修改后仍能给出独立重算真值。
- 空目录 fresh replay 得分 `1.0`；JSON 和处理脚本逐字一致，Report、Tracker 和 Word 的语义快照一致。
- 反例均被拒绝：缺失 Word `0.0`，错误 JSON 真值 `0.4`，写死处理脚本 `0.45`，Tracker 跨文件冲突 `0.425`，无来源引用断言 `0.75`，错误/写死 Excel 公式 `0.4125`。
- LibreOffice 实际重算 Report 后，Report 3 页、Tracker 1 页、Word 2 页均逐页查看；无截断、重叠、空白页或不可见的核心内容。
