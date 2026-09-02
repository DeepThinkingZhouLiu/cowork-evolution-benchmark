你是跨系统离职编排与内控分析员。请联合读取 `Offboarding_Master.xlsx`、`Initial_Application_State.json` 和 `Workflow_Context.json`，严格按冻结的本周日期边界判断人员范围，为范围内员工准备 BambooHR、Jira、Payroll 和 Tracker 的可审计离线动作包。

仅交付：`Offboarding_Multi_System_Audit.xlsx`、`Offboarding_Operations_Brief.docx`、`Offboarding_Operations_Review.pptx`、`BambooHR_Status_Update_Drafts.json`、`Jira_Access_Revocation_Drafts.json`、`Payroll_Notification_Draft.mbox`、`Tracker_Status_Update_Drafts.json`、`Offboarding_Execution_Control.json`。

- 本周范围固定为 2026-03-18 至 2026-03-22（含首尾）。Greg Foster / EMP-2001 / 2026-03-21 / PTO 5 days 与 Ian Costa / EMP-2004 / 2026-03-20 / PTO 0 days 在范围内；Nadia Petrova / EMP-2005 / 2026-04-05 不在范围内。
- Excel 必须有 `Source_Week`、`Source_Convention`、`Source_System_State`、`Evidence_Register`、`Scope_Decisions`、`BambooHR_Drafts`、`Jira_Drafts`、`Payroll_Draft`、`Tracker_Drafts`、`Skipped_Employees`、`Cross_System_Reconciliation`、`Control_Summary` 十二表；范围决策表须用公式逐格链接源表、计算日期边界、决策和目标状态。
- Greg 与 Ian 各有 1 条 BambooHR `Inactive` 草稿、1 条含姓名/员工ID/最后工作日的 Jira 撤权草稿、1 条 Tracker `Processed` 草稿，并共同进入唯一一封发给 `payroll@company.example.com` 的工资通知草稿；PTO、日期、类型和身份均须保真。
- Nadia 必须记录为 `SKIP_OUT_OF_SCOPE` 并保持 `Pending`；不得为她生成 BambooHR、Jira、Payroll 或 Tracker 动作。
- BambooHR 应用数、Jira 创建数、工资发送数、Tracker 应用数、外部动作数和真实副作用数均必须为 0；所有动作必须明确标记为未应用、未创建或未发送。
- Word 和三页 PPT 须解释本周范围、四系统一一对账、Nadia 跳过原因、PTO 保真和离线安全边界；PPT 使用原生可编辑表格。

不得真实修改 BambooHR、创建 Jira 工单、发送工资邮件或更新源 Tracker。
