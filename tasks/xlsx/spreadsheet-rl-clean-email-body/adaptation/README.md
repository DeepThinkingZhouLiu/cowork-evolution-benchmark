# spreadsheet-rl-clean-email-body

状态：`pending_human_review`；难度为 `L1` 作者假设，`calibration_status=uncalibrated`。

## 来源与证据

- 来源：Spreadsheet-RL，固定 revision `6f274cd4cbc27e656d7ddae7c3b591669866dc84`。
- 输入：`data/input_files/quotes_input.xlsx`，SHA256 `a6dc9cc59c40145f6aaa530835e31cacc49be91952016f6dfc91ce3e0a59815a`，角色为 solver-visible input。
- Golden：依据同一来源的公开目标工作簿制作并保存于隐藏侧 `solution/Golden/quotes_cleaned.xlsx`；Judge 逐格比较输入、目标和输出，检查目标变化与非目标保留。
- 许可：CC-BY-SA-4.0；当前按公开候选处理，最终发布仍需人工复核。

## 改造边界

保留原始工作流和输入，离线化运行环境，明确输出路径 `/app/output/quotes_cleaned.xlsx`，要求原生可编辑工作簿、公式/值保真和非目标结构保持。未加入 Golden 之外的业务事实。

## 验证记录

已运行 `PYTHONPATH=src python -m cowork_evo_bench.cli validate-task` 和 `validate-task-assets`；本题均通过。参考解法 fresh replay 与隐藏 Golden 均为 `1.000000`。空输出、损坏输出和未修改输入的 Judge 均不通过。Judge 覆盖文件完整性、目标一致性、公式/可编辑性、非目标保持、结构解析和输入哈希。尚未执行无辅助模型 rollout；难度和发布状态仍需人工复核。

## 已知限制

上游工作簿可能含 Excel 专有公式或图表行为；离线环境只验证 OOXML 结构和确定性目标，不替代人工打开、重算和视觉检查。当前候选不标记为 accepted 或 released。
