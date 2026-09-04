# R059 / Cowork-Bench financial-audit-cross-system-reconciliation

状态：`pending_human_review`；`difficulty_tier=expert` 是作者假设，当前 `uncalibrated`。

来源是 Cowork-Bench Apache-2.0 commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d` 的 `tasks/finalpool/financial-audit-cross-system-reconciliation`。原始 task/evaluator/config/license 保留在 `adaptation/`。三个输入逐字复制到 `data/input_files/` 和 `environment/input_files/`，SHA256 记录在 metadata：权威台账 CSV、手工登记 XLSX、系统匹配规则 PDF。

原任务输出 XLSX、DOCX、EML 和 ICS，并包含外部邮件/日历副作用。本适配依据 XLSX-only 契约，将自然交付环节压缩到一个七工作表 XLSX：源台账、手工表、对账、摘要、差异明细、邮件草稿和会议草稿。没有把非 XLSX 文件伪装为输出；邮件和会议均明确为未发送/未创建草稿。

Judge 从输入 CSV/XLSX/PDF 独立重算日期、金额、状态和缺失记录，不读取 Golden 或上游 expected JSON。固定真值为 15 条手工订单、11 条 matched、4 条 discrepancy（WC-005、WC-008、WC-011、WC-013），手工金额合计 3263.98。摘要公式经 LibreOffice 重算后检查，工作簿渲染为 12 页 PDF。

验证：`validate-task`、`validate-task-assets`、Golden、fresh replay 和完整渲染均通过，reward `1.000000`。负例计划覆盖空/损坏输出、错分状态、错摘要、遗漏差异记录和额外输出；当前尚未进行无辅助模型 rollout 或人工接受。leakage group 为 `cowork-financial-audit`，不得与同源对账任务跨切分。
