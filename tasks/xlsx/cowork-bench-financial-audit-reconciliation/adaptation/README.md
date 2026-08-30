# C093 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=advanced` 是 L3 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：Cowork-Bench commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`，`tasks/finalpool/financial-audit-cross-system-reconciliation`。
- 输入：`authoritative_ledger.csv`（SHA256 `6d141be371f2b2d8942e493b1f59b4a631ccd0eb482bb2ced6bcc3dd4989cb00`）、`q1_manual_tracking.xlsx`（`d980f32983abca1921f64282f9fb6d347e2080dcc7ac9ee721f7a8c43719d95e`）、`system_mapping_guide.pdf`（`457fd78134f2ea0877129436a13bb1cbce79270bf9f991d57a3139c03f61fddf`）。
- 许可：Cowork-Bench Apache-2.0；上游 task、evaluator、task config、expected output、两份 Golden 和 LICENSE 均保留在本目录。

## 独立真值

Judge 从 PDF 的四条规则和两份数据重新对账，不读取上游 expected JSON 或 Golden。手工表共有 15 个订单，总金额 3,263.98 RUB；11 个 matched，4 个 discrepancy：

| 订单 | 原因 |
|---|---|
| WC-005 | amount mismatch：手工 312.25，权威台账 310.00 |
| WC-008 | missing from authoritative ledger |
| WC-011 | date mismatch：2024-02-15 对 2024-02-20，超过 1 天 |
| WC-013 | amount mismatch：手工 145.75，权威台账 154.75 |

上游 expected output 和 evaluator 给出相同 15/11/4/3263.98 结论，仅作为交叉审计证据。

## 离线改造

原工作流从 InSales/ClickHouse 读取状态、实际发送邮件并创建 Google Calendar 事件。本包使用上游随题 `authoritative_ledger.csv` 作为冻结权威台账，保留 Excel 和 Word 输出，并将外部副作用替换为 `audit_notification.eml` 与 `audit_review_meeting.ics`。两者必须明确是离线草稿，不得声称已发送或已创建。未添加订单、人员或业务规则；会议具体时间不计业务真值，只检查有效且结束晚于开始。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 15 个逐单状态 | CSV、XLSX 和 PDF 匹配规则 | R002 |
| 15/11/4/3263.98 摘要 | 手工表和逐单重算 | R003 |
| Word 摘要和四项原因 | 同一独立对账记录 | R004/R005 |
| 财务负责人邮件草稿 | 原 task 的收件人、主题和摘要要求 | R006 |
| 审计复核会议草稿 | 原 task 的日历副作用 | R007 |
| 四个交付物一致 | 独立对账记录 | R008 |
| 来源与规则固定 | 三文件 SHA256 和 PDF 文本 | R010 |

## 验证记录

已执行 `validate-task`、`validate-task-assets` 和全库 `validate-public-tasks tasks`，均通过；全库当前为 65 个公开端到端任务。Golden 与从空目录运行 `solution/solve.sh` 的 fresh replay 均由标准 `tests/test.sh` 判为 `1.000000`，同一 Golden 连续 5 次均为 `1.000000`。

负例结果：空输出/损坏 XLSX/只缺 ICS=`0.000000`；把 WC-005 错分为 matched=`0.750000`；遗漏 WC-015=`0.750000`；Summary 总额错 1 RUB=`0.800000`；Word 差异原因错误=`0.850000`；邮件把 WC-013 写成 WC-014=`0.850000`。所有局部缺陷均因对应门槛失败而未通过，且展示了低于或等于阈值附近的评分梯度。Judge 从三份输入重算，不依赖本地 Golden 或上游 expected JSON。

LibreOffice 7.3.7.2 已打开并渲染 Excel 的两个工作表各 1 页、Word 1 页；15 行对账、四项摘要、四个差异原因均可见。EML 由标准库解析出收件人、主题和离线草稿正文；ICS 解析出单个 VEVENT、完整字段和有效时间顺序。人工预览未发现截断、重叠或空白内容。

## 难度与人工决策

- 三种输入、四个交付物、冲突处理和跨文件一致性构成长依赖链，作者暂定 L3；离线副作用是否足够保留原工作流需人工确认。
- leakage group 为 `cowork-financial-audit`；同一订单 fixture、映射 PDF 或对账语义变体不得跨切分。
- 最终难度、离线化充分性、许可结论和是否接受均待 rollout 与人工复核；当前没有可记为模型分数的无辅助运行。
