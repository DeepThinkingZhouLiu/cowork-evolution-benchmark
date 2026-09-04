# workspace-strategy-long-horizon（R053 / XLSX-C072）

状态：`pending_human_review`；作者难度假设 `L4`，`calibration_status=uncalibrated`。

## 来源审计

- Workspace-Bench-Lite 任务 `386`，metadata revision `60b08b1cc2e8054afbc3ca2160d37876b4f0765c`；角色 Product Manager。
- 实际 data_manifest 的 7 个输入均从 HF mirror 下载并保留 SHA256。Workspace-Bench-Workspaces revision `e245d63bfa20cfdb708cd8e78145ffb087155857` 的 README 声明 Apache-2.0，许可证据保留在 `LICENSE-Workspace-Bench-Workspaces`。
- metadata 还引用 D1 下午、D2 上午和音频文件，但这些不在 task 386 data_manifest，无法取得；`strategic_decision_evidence.json` 是基于 metadata key events/rubric 的最小离线证据投影，明确记录该边界。

## 改造边界与账本

原任务要求 PPTX、JSON、XLSX 三种输出；矩阵合同为 `one_xlsx`，因此以 7 个源文件和最小证据投影为输入，在一个可编辑 XLSX 中压缩决策确认、风险、OKR、指标、90 天行动计划和文件关系图谱。关系图保留原始 10 节点逻辑和 25 条依赖边，Traceability 行逐条列出风险及 OKR 来源。未提供的音频与转写不被伪造为输入。

## 验证

已运行 `validate-task` 与 `validate-task-assets`，均通过；Golden 与 fresh replay Judge reward 均 `1.000000`；空目录和损坏 XLSX 均以 reward `0` 拒绝。不得标记 `accepted` 或 `released`。
