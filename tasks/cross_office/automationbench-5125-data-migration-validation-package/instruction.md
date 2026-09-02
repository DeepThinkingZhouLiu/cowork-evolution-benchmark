你是员工数据迁移验证与 HRIS 更正治理分析员。请联合读取 `Employee_Data_Migration_Tracker.xlsx`、`Post_Migration_Corrections.mbox`、`HR_Slack_Context.json` 和 `Workflow_Context.json`，制作可审计、可回放的离线验证包。

必须以 `Employee ID` 而不是姓名匹配更正。两个 Michael Park 是不同员工：`EMP-6601` 的 Title 应由 Engineer 更正为 Senior Engineer；`EMP-6655` 在 Sales 的 Sales Rep 记录正确，不得报为需更正。Sarah Kim `EMP-6618` 的 Department 应由 Engineering 更正为 Product。邮件未指出更正的其余5条记录标记为 Validated。

仅交付：`Employee_Data_Migration_Audit.xlsx`、`Data_Migration_Validation_Brief.docx`、`Data_Migration_Validation_Review.pptx`、`HRIS_Correction_Draft.mbox`、`Validation_Status_Update_Drafts.json`、`Migration_Discrepancy_Register.json`、`Migration_Execution_Control.json`。

- Excel 必须有 `Source_Tracker`、`Source_Corrections`、`Evidence_Register`、`Validation_Decisions`、`Discrepancies`、`Validated_Records`、`Status_Update_Drafts`、`HRIS_Email_Draft`、`Reconciliation`、`Control_Summary` 十表；验证表须用公式链接源台账，依 ID 生成预期部门、预期职级、差异与状态更新草稿。
- 差异恰好2条：`EMP-6601` 的 Title，`EMP-6618` 的 Department。Validated 恰好6条，包含 `EMP-6655`。
- HRIS MBOX 恰好1封，发给 `hris@company.example.com`，只写两条更正和 `DRAFT_NOT_SENT`，不得把 `EMP-6655` 列为需更正。
- 状态更新 JSON 恰好8条，`EMP-6601`/`EMP-6618` 为 Discrepancy，其余为 Validated，全部 `DRAFT_NOT_APPLIED`、应用0。
- Execution 须记录审查8、差异2、验证6、状态草稿8/应用0、更正邮件草稿1/发送0、源行更新0、真实副作用0。
- Word 和三页 PPT 须解释两条差异、同名 ID 消歧、`EMP-6655` 正确记录和离线边界；PPT 使用原生可编辑表格。

不得真实发送邮件、修改源台账或应用状态更新。
