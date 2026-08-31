# sb2-debugging-03-03-lbo-repair

状态：`pending_human_review`；难度为 `L4` 作者假设，`calibration_status=uncalibrated`。

## 来源与证据

- 来源：SpreadsheetBench-2 sample，固定 revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece`。
- 输入：`data/input_files/03_03_input.xlsx`，SHA256 `429746c05f685b8de96b29cd29406ff85425952f60829b23dbdcef85682a1c72`，角色为 solver-visible input。
- Golden：依据同一来源的公开目标工作簿制作并保存于隐藏侧 `solution/Golden/lbo_03_03_repaired.xlsx`；Judge 逐格比较输入、目标和输出，检查目标变化与非目标保留。
- 许可：CC-BY-SA-4.0；当前按公开候选处理，最终发布仍需人工复核。

## 改造边界

保留原始工作流和输入，离线化运行环境，明确输出路径 `/app/output/lbo_03_03_repaired.xlsx`，要求原生可编辑工作簿、公式/值保真和非目标结构保持。未加入 Golden 之外的业务事实。

## 验证记录

已运行 `PYTHONPATH=src python -m cowork_evo_bench.cli validate-task` 和 `validate-task-assets`；本题均通过。参考解法 fresh replay 与隐藏 Golden 均为 `1.000000`。空输出、损坏输出和未修改输入的 Judge 均不通过。Judge 覆盖文件完整性、目标一致性、公式/可编辑性、非目标保持、结构解析和输入哈希。尚未执行无辅助模型 rollout；难度和发布状态仍需人工复核。

## 已知限制

上游工作簿可能含 Excel 专有公式或图表行为；离线环境只验证 OOXML 结构和确定性目标，不替代人工打开、重算和视觉检查。当前候选不标记为 accepted 或 released。
