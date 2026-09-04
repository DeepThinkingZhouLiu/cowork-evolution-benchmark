# 逐题改造记录

本题只改造 Cowork-Bench `terminal-canvas-excel-word-teamly-email`，不是按 `workspace_orchestration` 家族名称自编的任务。

- 上游仓库：`0717376/cowork_bench`
- 上游 commit：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 上游任务路径：`tasks/finalpool/terminal-canvas-excel-word-teamly-email`
- 原始 Canvas 快照：`db/init.sql.gz`，SHA256 `bd9c45b5ec9072b9f301b7df5b3e9ac65540ffcd53485329ebea44b0c3710da4`
- 许可：Apache-2.0

原题的工作链是：Canvas 两课数据 -> Python 风险评分 -> 四表 Excel -> Word 干预报告 -> Teamly 跟踪页 -> 学业顾问邮件。本地保留 SQL 快照的 2 条课程、4,648 条选课、26 条作业和 32,424 条提交源记录，只将数据库、Teamly 和邮件副作用静态化为 CSV、Markdown 和文本草稿。

`extract_canvas_snapshot.py` 是本单题的可审计投影脚本：它不生成、改写或补齐任何学生分数，只从上游 PostgreSQL COPY 段中保留课程 16/17 的必需列。

逐行重算真值：

- 课程 16：选课 `2,283`，已评分学生 `1,872`，非空提交 `16,228`，submission 均分 `77.14`，High/Medium/Low = `209/595/1,068`，pass rate `88.84%`。
- 课程 17：选课 `2,365`，已评分学生 `1,842`，非空提交 `16,175`，submission 均分 `78.72`，High/Medium/Low = `154/566/1,122`，pass rate `91.64%`。
- 合计已评分学生 `3,714`，High/Medium/Low = `363/1,161/2,190`，占比 `9.77%/31.26%/58.97%`。

上游冻结 Golden 把两课选课数写成 `2,285/2,369`，与当前 `db/init.sql.gz` 中的真实记录 `2,283/2,365` 冲突。本地不沿用这两个过期数字，Judge 直接从 CSV 重算。

## v2 连续评分闭环

- `rubric.v1.json` 与 `tests/judge_v1.py` 逐字节保留第一版严格口径；`SCORING_VERSION=v1|v2` 可显式回放，默认使用 v2。
- v2 只在评分器、JSON、Excel、Word、Markdown 或邮件缺失、为空、损坏/不可解析时硬归零；其他错误按两门课程、三档风险、三档干预、四张 Excel 表、三条公式、Word 章节、两行 Tracker、邮件字段和跨文件传播逐原子单元给连续分。
- 任务专属负例共 53 个：20 个缺失/空/损坏/原样输入致命反例均为 `0`；33 个单事实、阈值、公式、格式、跨文件冲突、整块缺失与严重但可解析反例落在 `0.463187–0.998841`。其中风险域跨文件冲突为 `0.775487`，风险域整体缺失为 `0.602821`，严重不完整但仍可解析的工作区为 `0.463187`。
- Golden、fresh replay 与 v1 Golden 均为 `1.0`。本题交付只有 Python/JSON/XLSX/DOCX/Markdown/TXT，没有 PPTX/PDF，因此不强加 VLM，最终分即 deterministic reward。
