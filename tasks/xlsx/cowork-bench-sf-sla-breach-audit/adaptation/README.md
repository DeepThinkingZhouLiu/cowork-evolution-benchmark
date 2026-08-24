# 逐题改造记录

本题只改造 Cowork-Bench `sf-sla-breach-audit`，不是从分类名生成的新故事。

- 上游仓库：`0717376/cowork_bench`
- 上游 commit：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 上游路径：`tasks/finalpool/sf-sla-breach-audit`
- 许可：Apache-2.0，完整文本保留为 `LICENSE-Cowork-Bench`
- 原始任务：俄文 Query，要求查询 ClickHouse 的工单与 SLA 策略、交付两张表的 Excel，然后发送结果邮件
- 本地改造：从同一上游 `db/init.sql.gz` 逐行抽取两张原始表；不再需要数据库或邮件服务；邮件动作改为 TXT 草稿
- 未改动：审计目标、关联键、违约定义、工作表名、列名、排序、摘要指标、收发件人和邮件主题

`extract_snapshot.py` 只复现输入的抽取过程，它不生成题目、Query 或答案。`upstream-groundtruth-SLA_Breach_Audit.xlsx` 只用于人工对照上游语义；本地 Judge 每次都从两份输入 CSV 重算真值，不读取该 Golden。

已人工确认上游 Golden 的核心结果与静态快照一致：31,588 条工单、24,961 次违约、总违约率 79.02%、最差优先级 `Medium`、最佳优先级 `Low`。
