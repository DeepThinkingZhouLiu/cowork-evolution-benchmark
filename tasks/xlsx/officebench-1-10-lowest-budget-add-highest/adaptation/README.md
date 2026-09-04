# officebench-1-10-lowest-budget-add-highest

状态：pending_human_review；difficulty_tier=basic 是 L1 作者假设，calibration_status=uncalibrated，review_status=pending_human_review。

## 来源与证据

- 来源：OfficeBench tasks/1-10/subtasks/3.json，固定 revision b978b808667c32b52ce19a67ce1def1de9ae02b7；原始子任务 JSON 与 Apache-2.0 许可证保留在本目录。
- 输入：data/input_files/company_budget.xlsx，SHA256 7b14e19c56e1b5185f61e16cda1f1b00d424d43cb41803ba5d7b9141a4a160a3，来自上游 testbed/data/company_budget.xlsx。
- 独立真值：Judge 从 solver-visible 输入重算目标，不读取 Golden；预期结果为输入驱动的单工作簿编辑。

## 改造边界

仅将上游任务改造为离线本地输入/输出，明确原生可编辑工作簿和非目标保持要求；未增加输入之外的业务事实。参考解法同样从输入计算并写出结果。

## 验证记录

作者侧应执行 validate-task、validate-task-assets、Golden/fresh replay，以及空输出、损坏输出、未修改输入和目标错误负例。Judge 覆盖文件完整性、输入驱动目标、非目标保持、OOXML 可解析性和输入哈希。尚未执行无辅助模型 rollout，最终难度和发布状态待人工复核。

实际验证（2026-08-31）：validate-task 与 validate-task-assets 通过；fresh replay/Judge=1.000000；空输出=-0.200000、损坏输出=-0.200000、未修改输入=0.350000。
