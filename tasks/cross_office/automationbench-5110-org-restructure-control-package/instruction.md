你是 HR Ops 的组织重组控制分析员。请联合审计 `/app/input_files` 中的员工目录、VP People 批准邮件、冻结运行时间和 HR Ops 频道，生成一套可复算、可继续编辑、可追溯，且不会真实发送邮件、更新目录或发布 Slack 的组织改线控制包。

只在 `/app/output` 中交付下列七个文件：

- `Org_Restructure_Audit.xlsx`
- `Org_Restructure_Control_Brief.docx`
- `Org_Restructure_Chart.pptx`
- `Org_Change_Notification_Drafts.mbox`
- `Directory_Update_Drafts.json`
- `HR_Ops_Reorg_Draft.json`
- `Scope_Control.json`

冻结处理时间为 `2026-04-15T09:00:00Z`，批准变更自 `2026-04-14` 生效。四条批准改线必须是：

- Aisha Jordan：`Lisa Wang -> Marcus Chen`，新经理邮箱 `marcus.chen@company.example.com`。
- Bo Lindqvist：`Marcus Chen -> Derek Huang`，新经理邮箱 `derek.huang@company.example.com`。
- Chiara Russo：`Derek Huang -> Lisa Wang`，新经理邮箱 `lisa.wang@company.example.com`。
- Dante Obi：`Lisa Wang -> Marcus Chen-Williams`，新经理邮箱 `marcus.chen-williams@company.example.com`。

新经理邮箱必须由员工目录中已有同名经理记录解析，不能猜测。`Marcus Chen` 与 `Marcus Chen-Williams` 是不同经理，Dante 绝不能错分给 Marcus Chen。Elena Vasquez、Felix Braun、Grace Kim 和 Hugo Fernandez 未出现在批准清单中，必须保持原经理和邮箱不变，不得进入更新或通知草稿。

用户提到的产品发布、办公室搬迁和 Engineering sprint teams 都是明确的无关事项。本题不得为它们生成任务、邮件、Slack、人员改线或其他动作；`Scope_Control.json` 必须逐项记录 `OUT_OF_SCOPE`、`action_count=0` 和来源片段。

Excel 必须至少包含 `Source_Directory`、`Approved_Changes`、`Manager_Lookup`、`Evidence_Register`、`Reorg_Decisions`、`Directory_Update_Queue`、`Notification_Queue` 和 `Control_Summary` 八张工作表。来源表按上游顺序完整保留八行；批准邮件全文、冻结时间和频道均进入证据表。`Reorg_Decisions` 的来源字段、是否受影响、旧经理核对、新经理、新经理邮箱、改线结果、通知、更新、Slack 摘要和追踪必须用原生公式派生；`Updated Manager` 对未受影响人员必须公式回落到原经理。队列与汇总也必须使用原生公式，并包含原生表格、条件格式和可编辑图表。

`Directory_Update_Drafts.json` 必须恰好包含 Aisha、Bo、Chiara、Dante 四条更新草稿，顶层和每行均为 `DRAFT_NOT_APPLIED`、`applied=false`，逐条保留源行、旧/新经理、新经理邮箱、生效日和来源追踪。

`Org_Change_Notification_Drafts.mbox` 必须恰好包含 4 封 `DRAFT_NOT_SENT`，分别只发给四名受影响员工；同一正文案例行必须包含员工、部门、旧经理、新经理、新经理邮箱、生效日和源行。不得有 Cc/Bcc，不得通知未受影响人员，也不得声称邮件或目录更新已执行。

`HR_Ops_Reorg_Draft.json` 必须引用 `C_HROPS_5110 / hr-ops`，为 `DRAFT_NOT_POSTED`、`posted=false`、真实消息数 0，并在一个待发布草稿中按批准顺序保留四条完整改线；不得加入四名未受影响员工或三个无关事项。

PPT 必须是 3 页、全部关键文字和组织图均可编辑：第 1 页为 4/4/0 的执行摘要和四条批准改线；第 2 页展示更新后的完整八人组织图，四个经理各有两名直属员工（Lisa：Chiara/Hugo；Marcus Chen：Aisha/Felix；Derek：Bo/Elena；Marcus Chen-Williams：Dante/Grace）；第 3 页展示经理身份消歧、四名未受影响人员和零真实执行边界。不得把组织图或表格做成整页截图。

Word 控制简报需覆盖八名员工、四条改线、四名未变人员、四个经理的解析证据、批准日、3 个无关事项、4 封通知、1 条 Slack 草稿和零真实执行。七个交付物中的人员、经理、邮箱、部门、顺序、计数、范围和执行边界必须一致。

所有 Office 文件必须能由 LibreOffice 无界面模式打开和渲染，不得使用宏、危险外链、调试文件或其他交付物。
