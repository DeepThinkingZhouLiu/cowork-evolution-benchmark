# 初始 Harness 目录

本目录保存可复现的初始 Harness 定义。H0 是 Benchmark 的固定条件，不属于某个 Evolution 算法的私有实现。

每个 H0 至少记录：

- 精确源码 Revision。
- 基础提示词、`AGENTS.md`、技能、记忆和工具说明。
- Agent 与模型启动方式。
- 依赖和办公工具版本。
- 单题资源预算和联网策略。
- 三条赛道各自的可写路径映射。

运行期间的 Baseline 副本和 Candidate 不保存在这里，而是写入 `.coworkevo/runs/`。
