你是 HR Ops 的入职设备配置控制分析员。请联合审计 `/app/input_files` 中的新员工台账、CFO 最新配置政策邮件、冻结运行时间和 IT Slack 频道快照，生成一套可复算、可继续编辑、可追溯且不会真实创建 Jira、发送消息或下采购单的 2026 年 4 月入职设备配置包。

只在 `/app/output` 中交付下列六个文件：

- `Equipment_Provisioning_Audit.xlsx`
- `Provisioning_Control_Brief.docx`
- `Provisioning_Jira_Drafts.json`
- `IT_Equipment_List_Draft.mbox`
- `IT_Channel_Draft.json`
- `Purchase_Execution_Control.json`

必须同时应用雇佣类型、已发货备注、部门和地点四层规则，不能机械地给台账中所有人建单：

- Ada Petrova：正式员工、Engineering、Austin 办公室；配置 `MacBook Pro 16-inch M4`、无线键盘、鼠标、耳机，不配额外显示器。
- Brian Osei：正式员工、Sales、Remote - Chicago；配置 `MacBook Air 15-inch M4`、标准三件外设和额外显示器。
- Carmen Delgado：`Contractor`；判为 `EXCLUDED_CONTRACTOR`，不得生成 Jira、邮件清单行或采购项。
- David Nguyen：正式员工、Engineering、Remote - Portland；配置 `MacBook Pro 16-inch M4`、标准三件外设和额外显示器。
- Elena Rossi：正式员工但备注写明设备已由原办公室发货；判为 `ALREADY_PROVISIONED`，不得重复生成 Jira、邮件清单行或采购项。
- Femi Adeyemi：`Independent Contractor`；判为 `EXCLUDED_CONTRACTOR`，不得生成 Jira、邮件清单行或采购项。
- Greta Holm：正式员工、Support、Austin 办公室；配置 `MacBook Air 15-inch M4`、标准三件外设，不配额外显示器。
- Hugo Fernandez：正式员工、Engineering、Remote - Madrid；配置 `MacBook Pro 16-inch M4`、标准三件外设和额外显示器。

政策优先级必须是：承包商/顾问排除和“设备已发货”排除先于部门、地点规则；其余正式员工中，Engineering 使用 Pro 16-inch M4，其他部门使用 Air 15-inch M4；所有实际待配置员工都有无线键盘、鼠标、耳机，仅 Remote 员工增加显示器。最终应有 5 个待配置、2 个承包商排除、1 个已配置；3 台 Pro、2 台 Air、3 台额外显示器。

Excel 必须至少包含 `Source_Starters`、`Policy_Rules`、`Evidence_Register`、`Provisioning_Decisions`、`Ticket_Queue` 和 `Control_Summary` 六张工作表。来源表按上游顺序完整保留原字段、源行号和原始字符串；政策表完整保留五项政策。证据表保留完整政策邮件、冻结时间、频道和八名人员的追踪。`Provisioning_Decisions` 中的来源字段、排除判断、最终状态、笔记本、显示器、三件外设、Jira 计划和邮件清单选择必须用原生公式从来源与政策推导；汇总也必须由原生公式计算。工作簿还要包含原生表格、条件格式和可编辑图表。

`Provisioning_Jira_Drafts.json` 必须恰好包含 Ada、Brian、David、Greta、Hugo 五个独立 issue 草稿，状态均为 `DRAFT_NOT_CREATED`、`created=false`。每个草稿必须包含人员、Employee ID、部门、地点、开始日期、正确笔记本、无线键盘、鼠标、耳机和是否配额外显示器，并引用来源行和政策。不得出现 Carmen、Elena 或 Femi，不得声称 Jira 已创建。

`IT_Equipment_List_Draft.mbox` 必须恰好包含一封只发给 `it-ops@company.example.com` 的 `DRAFT_NOT_SENT` 草稿，逐人列出上述五名待配置人员及完整设备；不得出现 Carmen、Elena 或 Femi，不得声称已经发送、下单或发货。

`IT_Channel_Draft.json` 必须引用 `C_IT_5107 / it-general`，状态为 `DRAFT_NOT_POSTED`、`posted=false`，只报告 8/5/2/1、3 台 Pro、2 台 Air、3 台额外显示器等聚合信息，不得披露姓名、Employee ID、地点或逐人设备。

`Purchase_Execution_Control.json` 必须记录 `DRAFT_ONLY_NO_ORDER`、`ordered=false`、实际采购记录数 0、实际 Jira 创建数 0，并说明 5 个 Jira 和 1 封邮件都只是草稿；不得包含订单号、采购交易、已创建 ticket key 或已发货状态。

Word 控制简报需覆盖八人、五项政策、5/2/1 状态分布、3/2 笔记本分布、3 台远程显示器、五个 Jira 草稿、一封 IT 邮件草稿、频道隐私摘要和零真实执行。Excel、Word、Jira JSON、MBOX、频道 JSON 和采购控制中的人员、型号、外设、排除、计数与执行边界必须一致。

所有 Office 文件必须能由 LibreOffice 无界面模式打开和渲染，不得用截图代替可编辑表格或图表，不得增加调试文件、危险外链、宏或其他交付物。
