# OfficeBench 3-5 逐题适配记录

本题直接读取上游 `tasks/3-5/subtasks/0.json` 和两份单页 PDF。两张来源都写 `Invoice #100`，但分别是 2023-01-30 的 Cloud service 与 2023-01-01 的 Decorative clay pottery；制题时没有创建第三张发票，也没有补写付款事件。

上游 evaluator 只检查 `invoices.xlsx`、`report.docx` 存在和 `cloud` / `pottery` 关键词。本适配将交付改为 5 页原生 PPTX，并用任务专属 Judge 核对两行发票登记、金额恒等式、原生堆积图、重复编号、29 天跨度、付款条款/状态边界和复核动作。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面变体均为 1.0。发票漏行/错金额、构成图错、汇总/日期错、编造付款状态/云供应商、pottery 物流错、缺来源、压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `4dc61c2e921d3c39fb409009ff9e47ba46b0bf6ef9f57666567d0e4dedcde6d8`。单题结构、资产、任务元数据、来源和 Rubric Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `officebench35-oracle-invoice-review-20260829`，trial 为 `officebench-3-5-invoice-payment__tWNTsbG`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 两份原发票和 Golden 5 页均已目视检查；表格、堆积图、付款边界和复核页无重叠裁切，`visual_qa.py` 返回空问题列表。
