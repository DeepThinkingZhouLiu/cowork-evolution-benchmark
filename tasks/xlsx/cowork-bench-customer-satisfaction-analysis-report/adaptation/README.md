# 逐题改造记录

本题仅改造 Cowork-Bench `customer-satisfaction-analytics-framework`，没有从任务分类名生成新数据或新故事。

- 上游仓库：`0717376/cowork_bench`
- 上游 commit：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 上游路径：`tasks/finalpool/customer-satisfaction-analytics-framework`
- 许可：Apache-2.0
- 原始输入：50 份问卷的 Excel，20 条客服工单的 CSV
- 原始交付：2 份 Excel + 2 份 Word

人工逐项读取输入后，可从来源直接复算：

- 问卷总体满意度 `4.20/5`
- NPS `42`（27 位推荐者、6 位贬损者）
- 客服工单首次解决率 `75%`（15/20）
- 客服工单平均解决时长 `3.40` 小时
- SMB 满意度 `4.00`，是三个分群中最低；SMB NPS 为 `-30`

上游 Golden 包含明显无来源内容：它把 50 份问卷写成 245 份，把 20 条工单写成 847 次服务互动，并附加了 `r=0.68`、“行业均值 30”、预算和收入影响。这些内容作为上游缺陷证据保留，不进入本地 Golden；本地 Judge 还会显式拒绝 245/847 样本量和上述无来源声明。

行动计划是原 Query 要求的开放交付。本地不对人名或某一句固定文案做匹配，只要求至少覆盖 FCR、解决时长和 SMB 三个由数据直接支持的改进方向，且基线正确、目标方向合理。
