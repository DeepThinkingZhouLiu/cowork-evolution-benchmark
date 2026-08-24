这道题逐题改造自 Cowork-Bench `rzd-msk-spb-team-trip-gcal-word`，不是本项目自编的差旅案例。

- 我实际阅读了上游俄文 Query、`trip_config.json`、`travel_guide.md`、preprocess、Evaluator、RZD MCP 查询实现、`db/init.sql.gz` 的车次/座席/路线记录，以及原始 `Travel_Plan.docx`。
- `trip_config.json` 和 `travel_guide.md` 原样复制；`travel_snapshot.json` 逐行投影自上游 SQL 和 preprocess 中的客户会议，不从本项目代码随机生成，也没有编造车次、时刻或票价。
- 原任务同时要求写日历、发邮件、保存 Markdown 和 Word。为形成责任单一的 `outline_to_document` 题，本改造只保留 Word 子任务，把动态服务读取冻结为离线快照；选择逻辑、预算约束、文档结构和原始 Golden 均保留。
- 上游数据本身是公开 Benchmark 的 mock RZD/日历环境，不是真实乘客订单，因此 provenance 标记 `synthetic=true`。

逐题验证记录：

- 原始 Golden 经独立 Judge 从三份输入重新推导后得分 `1.0`。
- `solve.py` 从输入回放得到的 7 个段落和三张表，与原始 Golden 逐项完全一致，回放得分 `1.0`。
- 空白、缺文件、错误去程、错误返程时间、错误舱位、错误总额、错误余额和缺章节共 8 个负例均未通过，最高负例得分 `0.658537`。
- LibreOffice 将 Golden 渲染为一页 PDF；人工查看确认标题、三节、三张表、车次、时间和金额均清晰，无截断、重叠或截图正文。

剩余风险：数据是上游构造的固定 2026 年 RZD 场景，不代表真实可售车票；正式 Release 仍需通过多 Harness rollout 确认难度与指令歧义。
