请基于以下 3 份冻结输入，制作一份 6 页、16:9 的三单匹配异常简报：

- `/app/input_files/Three_Way_Match_Batch.xlsx`
- `/app/input_files/Three_Way_Match_Tolerance_Policy.eml`
- `/app/input_files/AutomationBench_4044_Context.json`

必须满足：

- 通过 `PO #` 将 Invoice Batch、Purchase Orders 和 Receiving Reports 三张工作表的 4 笔记录逐条对齐，展示可追溯的原生可编辑来源表。
- 对每张发票展示 PO quantity、received quantity、invoice quantity、PO amount、invoice amount、amount delta、相对 PO 的 variance percentage 和最终状态。
- `VI-001` 的 `+$100.00 / +2.00%` 是允差边界，数量全部一致，状态为 `Approved for Payment`。
- `VI-002` 的 `+$250.00 / +7.14%` 是价格异常，上游断言状态为 `Price Variance - Review`。
- `VI-003` 的 PO 数量为 `30`，received / invoice 数量为 `25`，并且金额为 `−$1,400.00 / −14.58%`；数量规则优先，状态为 `Quantity Discrepancy - Hold`。
- `VI-004` 的数量与金额均精确一致，状态为 `Approved for Payment`。
- 展示 invoice batch `$21,450`、PO total `$22,500`、approved amount `$9,500`、exception amount `$11,950`，并用原生可编辑图表对比 4 笔 PO 与 invoice amount。
- 必须显式暴露一个上游语义缺口：政策只定义了价格偏差在 `within 5%` 时的 Review 状态，而 `VI-002` 为 `7.14%`；本任务依上游断言保留 Review，但不得声称政策已覆盖该区间。
- 对所有非批准项明确通知 `ap-lead@company.example.com`，并保留每页来源说明。
- 所有文本、表格和图表原生可编辑，无重叠、无裁切、无整页压平。

输出文件必须命名为 `/app/output/Three_Way_Match_Exception_Briefing.pptx`。请仅在 `/app/output` 中写入最终交付文件。
