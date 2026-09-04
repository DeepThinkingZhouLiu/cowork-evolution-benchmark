# cowork-inventory-optimization

状态：`pending_human_review`；作者难度假设为 `L4 / expert`，`calibration_status=uncalibrated`。本题是 Cowork-Bench 固定 revision `d943e75bc0fc8e3b27141979300cd8cbcd1e890d` 的逐题离线工作流重建，不是按 taxonomy 编写的合成故事。

## 来源与许可

- 上游任务：`tasks/finalpool/inventory-supply-chain-optimization`。
- 来源地址：<https://github.com/0717376/cowork_bench/tree/d943e75bc0fc8e3b27141979300cd8cbcd1e890d/tasks/finalpool/inventory-supply-chain-optimization>。
- 来源许可为 Apache-2.0，原文保留在 `LICENSE-Cowork-Bench`；上游题干、预处理器和 evaluator 分别保留在 `upstream-task.md`、`upstream-preprocess.py`、`upstream-evaluator.py`。
- 该快照是公开 benchmark mock fixture，不含生产个人数据。所有 solver-visible 输入及 SHA256 见 `metadata/input-manifest.json` 和 `metadata/provenance.json`。

## 来源审计与改造边界

上游题干依赖 InSales、ClickHouse、供应商 HTTP API、Google Sheet、Google Calendar、邮件和 Word。固定 commit 中的公开 mock 数据被投影为五个最小输入：商品、订单、供应商目录、类别需求和快照日期。产品/订单字段只保留本题所需的库存、名称、类别、状态、日期、商品明细和数量；客户、地址、电话、邮箱等无关字段不进入输入。

原始外部写操作全部去除：Google Sheet、日历、邮件和 Word 不被伪装为已经执行，改为同一 XLSX 中可复算的 `Restock_Plan` 和 `Restock_Summary`。类别需求从 `store_orders.json` 的有效订单和 `store_products.json` 的商品类别独立重算，再与冻结的 `demand_history.csv` 对照。没有从上游 Golden 反推任何 solver-visible 输入；上游 `expected_output.json` 仅记录 `records_processed=0`，不作为真值。

## 证据到需求映射

| 需求 | 证据 | Judge 方法 |
| --- | --- | --- |
| 商品筛选、库存和名称 | `store_products.json` | 重新筛选 `publish` 且正库存，核对 77 个 ID |
| 有效销量和每日速率 | `store_orders.json`、任务规则 | 重新筛选 `completed/processing`，汇总 line items，核对 `Total_Sales/90` |
| 剩余天数和补货标识 | 任务规则 | 读取重算后的公式值，核对 999、1 位小数及 `<30` |
| 供应商事实与风险 | `supplier_catalog.json`、任务规则 | 核对四个供应商字段及风险指数/风险带公式 |
| 类别历史需求 | `demand_history.csv`，并由商品/订单重算 | 逐类别核对总量和订单数 |
| 补货优先级与摘要 | `Stock_Status`、任务规则 | 核对补货集合、排序、跨表链接和摘要公式值 |
| 可编辑性与可回放性 | XLSX OOXML、LibreOffice | 检查表格/冻结窗格/自动重算并执行重存 |

## 固定快照复算

- 82 个商品中 77 个符合 `publish` 且正库存筛选。
- 150 个订单中只计 `completed` 和 `processing`；销量合计按产品明细重新汇总。
- 6 个类别需求行：Аудио 86/43、Бытовая техника 44/24、Камеры 67/31、ТВ и домашний кинотеатр 50/27、Часы 20/10、Электроника 160/78。
- `Stock_Status` 中两项剩余天数低于 30：产品 ID 22 和 31；平均剩余天数按显示值为 399.4。
- 供应商目录为 RUB，四个供应商字段完整保留；风险指数只由交付期和可靠性计算，不推断产品与供应商的未提供映射。

## 验证与交接

参考解法通过 `solution/solve.sh` 生成 `Inventory_Optimization_Report.xlsx`，Golden 位于 `solution/Golden/`。2026-09-04 的验证记录如下：

```text
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/xlsx/cowork-inventory-optimization
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/xlsx/cowork-inventory-optimization
```

以上两条命令均输出“校验通过”。Golden Judge 和独立 fresh replay 均为 reward `1.000000`、11/11 判据通过；LibreOffice 重算、PDF 渲染和 OOXML 结构检查均通过。空目录、损坏 XLSX 的 Judge reward 为 `0.000000`；删除一行和把摘要公式改成硬编码分别被拒绝（reward `0.333333`、`0.700000`）。

Judge 只从输入和显式规则重算，不接受 Golden 文件哈希或固定布局作为答案。当前未执行无辅助模型 rollout，不能把该题标记为 `accepted` 或 `released`。

## 人工复核事项

1. 复核将外部供应商 API、ClickHouse 和副作用压缩为静态 XLSX 的离线化边界是否符合最终 benchmark 发行策略。
2. 复核 `Risk_Index` 与 `Risk_Band` 的阈值是否应继续作为明确的适配规则，或改为仅展示指数。
3. 结合模型试做失败模式校准 L4 难度，并检查俄文商品/类别名称在目标 Excel 环境中的可读性。
