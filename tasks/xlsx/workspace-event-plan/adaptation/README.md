# workspace-event-plan（R052 / XLSX-C067）

状态：`pending_human_review`；作者难度假设 `L4`；`calibration_status=uncalibrated`。

## 来源与许可

- 任务元数据来源：`https://huggingface.co/datasets/Workspace-Bench/Workspace-Bench-Lite`，revision `60b08b1cc2e8054afbc3ca2160d37876b4f0765c`，`task_lite_clean_cn/159/metadata.json`。该元数据仓未声明许可证，仅作为任务意图、字段和依赖清单证据，不把其自身作为二进制再分发授权。
- 输入文件来源：`https://huggingface.co/datasets/Workspace-Bench/Workspace-Bench-Workspaces`，revision `e245d63bfa20cfdb708cd8e78145ffb087155857`，Apache-2.0（许可证文本保存在 `LICENSE-Workspace-Bench-Workspaces`）。按 Lite manifest 取出 60 个 Markdown 与 `event_plan_v1.xlsx`；哈希见 `metadata/input-manifest.json`。
- 上游输入为公开 benchmark fixture，按保守口径标记 `synthetic=true`；未发现个人敏感数据。

## 改造与证据账本

| 要求/事实 | 证据 | 改造/验证 |
|---|---|---|
| 60 条活动、Event 1–60 | 60 个 `event_plan_N.md` | `solve.py` 逐文件解析并按 Event ID 排序；Judge 重新解析 |
| 十个输出列 | metadata rubric 与 Markdown Event Details/Team Assignments | Golden `Event Plan` 原生表格，Judge 检查表头 |
| v1 存在错位/缺失 | `event_plan_v1.xlsx`（例如 Location/标题错位）与各 Markdown 对照 | 不复制 v1 的值；仅使用其作为编辑壳，所有字段来自可见 Markdown |
| 日期、人数、线索目标、负责人和状态 | 每个 Markdown 的 Event Details、Objectives、Team Assignments、Status | Judge 对 60 行逐字段重算，日期要求为真实日期值 |

## 验证与交接

- Golden：`solution/Golden/event_plan_v2.xlsx`；参考解法从输入重新生成，不读取 Golden。
- 独立 Judge：`tests/judge.py`；Golden 得分 `1.000000`。
- 已覆盖：缺失/损坏输出拒绝、行数/连续编号、十列表头、逐行 Markdown 一致性、日期/整数/命名格式、额外文件检查。
- 仍需人工复核：上游任务原本强调“读取并使用所有 Markdown”，本适配将外部工作区探索压缩为显式输入；正式发布前需确认该投影与许可证边界、难度和泄漏分组。
