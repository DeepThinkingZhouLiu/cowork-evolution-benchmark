# 逐题改造记录

## v2 连续评分改造（2026-09-04）

- 原严格评分逐字节归档为 `rubric.v1.json` 与 `tests/judge_v1.py`，SHA256 分别为 `d1b74bc7ad766020739fa04c30bf6854bd4c47df1e6ece9290359a4cfccda72c`、`1a868d51d47ca0f95039e4def80def82b1e0850c9af539f63fabe0987dd372f8`；原题没有独立反例脚本可归档。默认使用 v2，设置 `SCORING_VERSION=v1` 可复现旧口径。
- 本题不是把“大项是否全对”直接映射成 0/1，而是逐项检查脚本的原始/变异回放、JSON 字段、Report 行与公式、Tracker 行、俄文 Word 的结构与事实、来源边界、跨文件一致性、原生可编辑格式和 LibreOffice 渲染。
- 只有 5 个核心输出缺失、为空、损坏或不可解析才硬归零；单个 Paper_ID、增长率、公式、状态、标题或格式错误只失去对应原子分。`quant-ph` 噪声越界、脚本写死、额外输出和无来源断言使用互不重复的任务专属分级处罚。
- v2 回归包含 17 个致命反例和 19 个局部反例：致命反例全部为 `0`，局部反例分布为 `0.742500–0.998684`。Golden、fresh replay 和 v1 Golden 均为 `1.0`，空目录与仅输入文件均为 `0`。
- 本题只交付 JSON、Python、XLSX 与 DOCX，没有 PPTX/PDF 视觉交付，因此按统一协议不调用 VLM；审美与可读性由可编辑结构、格式原子检查及 LibreOffice 实际渲染承担。

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

- v1 Golden 与 v2 Golden 均为 `1.0`；v2 处理脚本在 Paper_ID 和两个 Growth_Pct 被修改后仍能给出独立重算真值。
- 从空输出目录运行 `solution/solve.py` 的 fresh replay 为 `1.0`；空交付与只复制输入文件的交付均为 `0.0`。
- v2 的 17 个缺失/空/损坏/输入原样反例全部为 `0`；19 个单事实、单公式、单标题、单格式、脚本响应和额外文件反例均得到可区分的连续分。
- LibreOffice 实际重算 Report 后，Report 3 页、Tracker 1 页、Word 2 页均可打开并出现核心文字；本题没有适用的 PPTX/PDF VLM 评分。
