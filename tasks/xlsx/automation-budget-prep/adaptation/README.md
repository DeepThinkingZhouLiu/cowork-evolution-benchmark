# AutomationBench finance.annual_budget_prep 逐题改造记录

状态：`pending_human_review`；`difficulty_tier=advanced` 是作者的 L3 假设，尚未校准或发布。

## 来源、许可和输入

- 来源仓库：`zapier/AutomationBench`，固定 commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，任务函数 `get_fin_annual_budget_prep_task`（`automationbench/domains/finance/tasks.py`）。
- 上游许可为 MIT；完整 `LICENSE` 与上游 finance 任务定义保存在本目录。状态中的部门和金额是仓库明确的 mock fixture。
- `budget_preparation_source.xlsx` 是对上游 `ss_budget_prep` 的 Rate Card、2025 Actuals 和空白 2026 Budget 的逐字段离线投影；输入 SHA256 为 `e8437abfd82569d54761cbe45f348ae66013a61def0272cf780db450983f4f03`。

## 真值与改造边界

Rate Card 给出的类别增长率分别为 Salaries 4%、Travel 10%、Software 20%。Include 部门 Engineering、Sales、Marketing 的结果分别是 (1,560,000; 44,000; 240,000)、(832,000; 165,000; 60,000)、(520,000; 66,000; 120,000)，总计 (2,912,000; 275,000; 420,000)。LegacyOps 的 Scope 为 Exclude，不能写入 2026 Budget 或邮件正文。上述事实均从输入独立重算；未使用上游 Golden 反推。

原题的 Google Sheets 追加动作在输出工作簿中保留为 `2026 Budget` 表；每个金额使用引用 `2025 Actuals` 和 `Rate Card` 的公式，末行 `Total` 使用 SUM。原题 Gmail 发送动作改为同一工作簿内明确标记 `Draft - Not Sent` 的 `CFO Email Draft`，不声称外部发送；由于原 prompt 要求跳过 Exclude 部门，草稿正文不列出 LegacyOps。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 部门集合、Scope 排除 | `2025 Actuals` 输入和原 prompt/assertions | R002/R004 |
| 类别增长率及金额 | `Rate Card` 与 2025 金额输入独立计算 | R003 |
| 公式和 Total | 跨表业务公式和用户可审计要求 | R004 |
| CFO 邮件草稿 | 原 prompt/assertions 的收件人与汇总语义 | R005 |

## 验证与风险

参考解法和 Judge 从输入重算并检查公式，未读取 Golden；空输出、损坏工作簿、统一增长率、LegacyOps 误入和金额篡改负例应失败。应执行 `validate-task`、`validate-task-assets`、Golden/fresh replay、LibreOffice 重存以及改变 Rate Card 的可重算测试。发件人、邮件正文是否保留 LegacyOps 排除说明及 L3 难度由人工复核。
