# C008 逐题改造记录

状态：`pending_human_review`。这是 pilot 的 L2 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench，commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，任务 `tasks/1-6/subtasks/2.json`。
- 原始 Query：`assign each students in Class 1 10 min ... start from 8:00, record the start time under new header 'Meeting'`。
- 原 evaluator 检查 `B1=Meeting`、`B2:B6=8:00/8:10/8:20/8:30/8:40`。
- 输入：`class_member.xlsx`，SHA256 `7d62695b7c450cbe61059bf7e32e7dd0f901718e8f5de1e6fd34714123e2e85e`，6,320 bytes。
- 许可：上游仓库 Apache-2.0；原任务记录和 LICENSE 由作者侧保留。

## 改造边界

输入只有 `Name` 列，没有独立的班级列。依据原 Query 和 evaluator，将输入中的五行（含重复 Noah）解释为该任务的完整 Class 1 名单，明确要求按输入行顺序填写，不引入额外筛选。时间按上游断言保留为 `8:00` 到 `8:40` 字符串。只增加 B 列，不改变 A 列。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| Meeting 表头与时间序列 | 原 Query、上游 evaluator | R002 |
| 姓名顺序（含重复项）保持 | 输入工作簿 | R003/R004 |
| 原生可编辑工作簿 | 输入 OOXML、保存后可读取 | R001/R006 |

## 验证记录

参考解法从输入逐行写入时间；Judge 独立检查表头、序列、姓名和非目标变化。已执行：

- `validate-task`、`validate-task-assets`：通过。
- Golden 与 fresh replay：Judge `1.000000`。
- 空输出 `0.000000`、损坏 XLSX `0.000000`、未修改输入 `0.500000`（结构和核心结果门槛失败）。
- LibreOffice 7.3.7.2 打开并导出 PDF，文本包含 `Meeting`、`8:00` 和 `8:40`。

## 风险与人工决策

- “Class 1” 只存在于原 Query，不在输入字段中；当前解释是由上游 evaluator 支持的最小澄清，需人工确认。
- 同源其他课程/排程题需要按 leakage group 隔离。
- Codex/OpenCode rollout 尚未执行；`difficulty_tier=intermediate` 为 `uncalibrated` 假设。
