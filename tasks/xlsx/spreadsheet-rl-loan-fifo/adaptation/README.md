# Spreadsheet-RL FIFO 改造记录

来源为 Spreadsheet-RL 固定 revision `6f274cd4cbc27e656d7ddae7c3b591669866dc84` 的 ExcelForum task-job `b07dec0d38d63663`。输入工作簿从归档提取为 `transactions_input.xlsx`，SHA256 为 `3195814da1d869b35721a1a60f100192b5ddd2e106971c9856a0b0a33def2fa2`；上游 instruction 指定 `FIFO_Alloc!AB4:AB10`。

参考解法将七笔 SELL 与此前 BUY 批次建立可编辑 FIFO 分配公式；Judge 从 Transactions 独立按顺序耗用批次，重算七个加权平均成本（94.00、94.375、94.80、95.00、96.00、96.14、96.70 c/$），并检查原表保留。作者假设为 L3，状态保持 `uncalibrated`、`pending_human_review`。
