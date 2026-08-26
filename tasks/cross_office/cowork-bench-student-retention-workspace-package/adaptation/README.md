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
