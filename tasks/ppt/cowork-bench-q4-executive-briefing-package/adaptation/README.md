本题逐题改造自 Cowork-Bench 的 `ppt-clickhouse-executive`，精确来源版本为 `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`。

- `upstream-task.md`：上游俄文用户 Query，原样保留。

- `upstream-task-config.json`：上游需要的工具和服务声明，原样保留。

- `upstream-evaluator.py`：上游从数据库动态查询真值的评分器，原样保留。

- `upstream-groundtruth-*`：上游原始 Excel、PPTX 和 PDF 参考交付物，原样保留，不作为本地 Judge 的真值来源。

- `LICENSE-Cowork-Bench`：上游 Apache-2.0 许可证。

- `extract_snapshot.py`：只处理本题指定的两张表，从上游 `db/init.sql.gz` 抽取 Q4 2024 订单和有关客户维度。这是离线化脚本，不是通用制题器。

原题的核心任务是：连接订单和客户信息，按地区和客户分群聚合 Q4 2024 销售，再交付 Excel 分析、高管 PPT 和同版 PDF。本地改造只将 ClickHouse 动态依赖替换为两份静态 CSV，没有自编新的业务故事、数据或交付要求。

输入快照来自 Cowork-Bench 的模拟企业数据库，因此来源元数据中 `synthetic` 保持为真。快照只保留订单编号、客户编号、日期、金额、分群和地区，没有导出姓名、邮箱等不必要字段。

当前上游 SQL dump 的地区和分群标签已集中替换为俄语，但上游仓库内的旧 GroundTruth 仍使用英文标签。数值聚合一致，标签不一致。为避免把仓库内部版本偏差写进 Benchmark，本题 Golden 和 Judge 都以当前精确 commit 的 SQL dump 快照为准；Judge 每次从输入 CSV 重算真值，不与本地 Golden 做字节比较。

参考交付物是针对这一道上游真实任务逐页制作的。Excel 保留原题指定的两张表、表头、排序和计算口径；PPTX 为 7 页可编辑 16:9 高管简报；PDF 由该 PPTX 实际导出。`solve.py` 只复现这一题的参考改造，不生成任何其他题目。
