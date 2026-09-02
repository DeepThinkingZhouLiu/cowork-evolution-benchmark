# AutomationBench 5070 逐题改造记录

- 上游：`get_hr_offboarding_checklist_multi_app_task()` / `hr.offboarding_multi_system`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `b48d77da6fef11c2c69cd5c821f22c9dec3526d696c827fce1dd96dbc6cc6217`。
- 已完整读取原 Query、3 人 Tracker、BambooHR/Jira 初始状态和 9 条断言；`materialize_inputs.py` 只处理 5070。因 XLSX 工作表名最长 31 字符，输入表把上游 `Offboarding Processing Convention` 缩短为 `Offboarding Convention`，原文保留在精确快照。
- 人工真值：本周为 2026-03-18 至 2026-03-22，Greg 与 Ian 需进入 BambooHR/Jira/Payroll/Tracker 四系统离线草稿；Nadia 的最后工作日为 2026-04-05，必须跳过且保持 Pending。
- Golden 含 12 张 Excel 表、36 个公式、3 页 Word、3 页 PPT、6 条系统动作 JSON、1 封 MBOX 草稿和 1 份控制 JSON。Judge 独立核验日期边界、身份、PTO、Processed 约定、四系统对账、Nadia 排除、零执行与 Office 安全。
- 13 类负例全部 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；18 页目视检查通过；两次 Harbor Docker Oracle 回放均为 1.0、0 异常。
