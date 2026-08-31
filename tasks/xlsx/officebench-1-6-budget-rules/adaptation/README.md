# officebench-1-6-budget-rules

状态：`pending_human_review`；难度为 `L2` 作者假设，`calibration_status=uncalibrated`。

## 来源与证据

- 来源：OfficeBench，固定 revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`。
- 输入：`data/input_files/budget_input.xlsx`，SHA256 `7b14e19c56e1b5185f61e16cda1f1b00d424d43cb41803ba5d7b9141a4a160a3`，角色为 solver-visible input。
- Golden：依据同一来源的公开目标工作簿制作并保存于隐藏侧 `solution/Golden/budget_rules.xlsx`；Judge 逐格比较输入、目标和输出，检查目标变化与非目标保留。
- 许可：Apache-2.0；当前按公开候选处理，最终发布仍需人工复核。

## 改造边界

保留原始工作流和输入，离线化运行环境，明确输出路径 `/app/output/budget_rules.xlsx`，要求原生可编辑工作簿、公式/值保真和非目标结构保持。未加入 Golden 之外的业务事实。

## 验证记录

已运行 `PYTHONPATH=src python -m cowork_evo_bench.cli validate-task` 和 `validate-task-assets`；本题均通过。参考解法 fresh replay 与隐藏 Golden 均为 `1.000000`。空输出、损坏输出和未修改输入的 Judge 均不通过。Judge 覆盖文件完整性、目标一致性、公式/可编辑性、非目标保持、结构解析和输入哈希。尚未执行无辅助模型 rollout；难度和发布状态仍需人工复核。

## 已知限制

上游工作簿可能含 Excel 专有公式或图表行为；离线环境只验证 OOXML 结构和确定性目标，不替代人工打开、重算和视觉检查。当前候选不标记为 accepted 或 released。
