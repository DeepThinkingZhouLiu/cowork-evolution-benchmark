# workspace-bank-ledger（R036 / XLSX-C069）

状态：`pending_human_review`；作者难度假设 `L3`；`calibration_status=uncalibrated`。

## 来源与许可

- 任务意图与 rubric：Workspace-Bench-Lite task 255，metadata revision `60b08b1cc2e8054afbc3ca2160d37876b4f0765c`。
- 输入：Workspace-Bench-Workspaces revision `e245d63bfa20cfdb708cd8e78145ffb087155857`，HF 数据卡声明 Apache-2.0；许可证文本保存在 `LICENSE-Workspace-Bench-Workspaces`。输入是 `LogisticsManager_Workdir/财务分析/原始数据/银行存款日记账_完整版.xlsx`，哈希见 `metadata/input-manifest.json`。

## 关键冲突与处理

Lite metadata 的若干 rubric 给出了另一版账本的金额（例如 1 月借方 704,587.92、年度借方 20,067,765.53），但 revision `e245d63...` 中实际可见 XLSX 的合计行和交易明细为 1 月借方 234,862.64、年度借方 2,517,952.91。依据 skill 的冲突顺序，本适配以 solver 可见输入为准，独立 Judge 从交易明细重算；未将 metadata 数字硬编码为真值。该差异是人工复核风险，题包必须保持 `pending_human_review`。

## 改造与验证

- 原任务的 XLSX 与 Markdown 双交付被压缩为单一 `月度资金流动趋势表.xlsx`，其中包含月度资金趋势、年度汇总、账户基础信息、高频交易对手和“分析报告”工作表。
- Golden `solution/Golden/月度资金流动趋势表.xlsx` 由 `solution/solve.py` 从账本交易明细重算生成，不复制上游 Golden。
- `tests/judge.py` 独立读取输入，检查 12 月度行、月度/年度借贷、余额、目标完成率、交易对手频次、章节与建议；Golden 和 fresh replay 均应在 LibreOffice 重算后得分 `1.000000`。
- 未决：上游 rubric 与当前 revision 的金额冲突需人工决定是否保留此候选；正式发布前还需渲染检查和难度 rollout。
