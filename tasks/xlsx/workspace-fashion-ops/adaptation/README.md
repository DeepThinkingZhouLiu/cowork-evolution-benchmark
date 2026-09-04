# workspace-fashion-ops（R037 / XLSX-C070）

状态：`pending_human_review`；作者难度假设 `L3`，`calibration_status=uncalibrated`。

## 来源审计

- Workspace-Bench-Lite 任务 `277`，固定 metadata revision `60b08b1cc2e8054afbc3ca2160d37876b4f0765c`；角色 Operations Manager。
- 输入原始文件：`9372683673dd8dc1_服装销售管理系统.xlsx`，由 `task_lite_clean_cn/277/data/` 下载，SHA256 记录于 `metadata/input-manifest.json`。
- 归档来源为 Workspace-Bench-Workspaces revision `e245d63bfa20cfdb708cd8e78145ffb087155857`，仓库 README 声明 Apache-2.0；许可证文本保留在 `LICENSE-Workspace-Bench-Workspaces`。

## 改造与证据账本

原始任务要求 Markdown 运营报告和 Excel 汇总表；矩阵输出合同为 `one_xlsx`，因此将报告中的可验证摘要压缩到同一可编辑 XLSX。销售报表实际含 66 条完整记录（总额 1,233,140 元、数量 4,285、10 名销售人员），而导购业绩与进货明细尾部出现“其余记录结构同上”占位符。适配只使用可见记录，不补造省略数据；综合毛利率 29.81%仅作为上游 metadata 断言并明确限制。

## 验证

已运行 `validate-task` 与 `validate-task-assets`，均通过；Golden 与 fresh replay Judge reward 均 `1.000000`；空目录和损坏 XLSX 均以 reward `0` 拒绝。所有题包只能保持 `pending_human_review`，不得标记 `accepted` 或 `released`。
