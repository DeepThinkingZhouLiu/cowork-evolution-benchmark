# 按 FIFO 计算贷款 SELL 交易成本

请在 `transactions_input.xlsx` 的 `Transactions` 工作表基础上，新增 `FIFO_Alloc` 工作表。对每一笔 `SELL` 按交易发生顺序使用此前尚未耗用的 `BUY` 批次进行 FIFO 分配，并在 `Avg Cost (c/$)` 列中计算该笔 SELL 所释放数量的加权平均成本（单位为 cents per dollar）。

`FIFO_Alloc` 应列出每个 SELL 一行、每个 BUY 批次一列分配数量，并在 `Relieved Cost ($)` 和 `Avg Cost (c/$)` 中使用可重算公式；不要把结果硬编码。保留原 `Transactions` 工作表。将文件保存为 `/app/output/transactions_fifo.xlsx`。
