# SpreadsheetBench-2 12_01 改造记录

来源为 `icyCreater/SpreadsheetBench-2-sample` 固定数据集 revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece` 的 `Financial_Model/12_01`；任务记录和评测代码来自 SpreadsheetBench-2 固定提交 `5c160265aa93c15b38e4034cbf1e09ab498335d9`。数据集许可证为 CC-BY-SA-4.0，全文保存在 `LICENSE-CC-BY-SA-4.0.txt`。输入 `12_01_input.xlsx` 已固定为 `financial_model_input.xlsx`，SHA256 为 `7f258ae78f7b8b7856a56b91a3a30632d1b11f02e4a50dd59be2d80e89fb42fd`；上游 Golden 作为审计证据保存在 `upstream-golden.xlsx`，不被 Judge 读取。

适配保留原任务边界：Revenue 按 Y-o-Y Growth，COGS 按 Total Revenue 百分比，Receivable 按 70 天，NWC 变动按逐期滚动，EBITDA Margin 按 EBITDA/Total Revenue。参考解法只写目标空白单元格并打开自动重算；Judge 从输入独立构造目标公式，逐格检查非目标保持，并尝试用 LibreOffice 重算。作者假设为 L3，保持 `uncalibrated`、`pending_human_review`。
