# AutomationBench 1115 逐题适配记录

本题直接阅读了 AutomationBench `marketing.competitive_response_plan` 的原 Prompt、完整 initial state 和 assertions，没有用通用生成器批量编题。上游 Google Sheets 的 6 行 feature、分析师邮件、价格政策邮件和 3 条 Slack 消息被分别静态化，以保留源边界。

Golden 是人工逐页设计的 6 页 16:9 竞争响应简报：封面结论、原生 gap-status 图表、6 行原生 feature 矩阵、分析师/Slack 分源证据、响应优先级和价格治理/交付页。不执行外部邮件、HubSpot 或价格页写入，但把原断言中的 brief 事实、ticket 主题和禁止直改要求转为可确定检查的演示交付。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面配色变体均为 1.0。图表分布错误、矩阵漏行/gap 错、分析师或 Slack 事实错误、响应优先级错误、价格治理违规、受众/ticket 错、压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `f4ef3ab496a253a8d57e458beb3b42ec80f1ff69cfa998b0498fd21aee9ab49b`。单题结构、公开资产、任务元数据、来源和 Rubric Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `automationbench1115-oracle-competitive-20260829`，trial 为 `automationbench-1115-stratacorp__JRwvF3P`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 6 页已渲染为接触表并逐页目视；原生 gap 图、6 行 feature 矩阵、双源证据、响应和治理页均无重叠裁切。`visual_qa.py` 返回空问题列表。
