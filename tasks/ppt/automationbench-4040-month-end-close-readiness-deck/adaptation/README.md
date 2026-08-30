# AutomationBench 4040 逐题适配记录

本题直接阅读了 AutomationBench `finance.financial_close_checklist` 的原 Prompt、工具列表、完整 initial state 和 7 条 assertions，没有用通用题目生成器。

上游任务原本要求 Agent 从 Google Sheets 读取 7 行 January close checklist，向 Controller/CFO 发邮件并向 `finance-alerts` 发 Slack。本 PPT 适配不执行外部写入，而是把上游应用状态静态化成一个 XLSX 和一个 JSON，要求 Agent 生成同等事实范围的管理简报。输入中的 7 个 item、姓名、邮箱、Status、Due 和唯一 Notes 均逐字段来自上游；未编造金额或额外财务事实。

Golden 是根据这两个冻结输入逐页设计的 5 页 16:9 演示文稿：封面/总体结论、原生完成度图表、7 行原生明细表、3 个未完成责任卡和面向 Controller/CFO/Slack 的沟通绪论。

## 实际验证

- 本地 Golden、ZIP 等价重封装和非语义性封面配色变体均为 1.0。统计错、图表数值错/缺失、明细漏项/负责人错、Tax 备注漏失、due 分类错、受众/频道错、编造金额、压平和答案标记均为 0；多一个调试文件为 0.94。
- Judge 还会拒绝同一文稿中并存 57.1% 和错误完成率的相互矛盾交付。fresh replay 连续 5 次均为 1.0；单题结构和输入/Golden 资产校验通过。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `automationbench4040-oracle-close-readiness-20260829`，trial 为 `automationbench-4040-month-end-c__n7r9PeJ`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 5 页已渲染为接触表并逐页目视；原生环图、7 行明细表、3 张责任卡和沟通末页均无重叠裁切。`visual_qa.py` 返回空问题列表。
