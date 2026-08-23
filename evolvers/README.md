# Evolution 算法目录

本目录保存符合统一接口的 Evolution 算法和论文参考基线。

算法输入是标准 Feedback Packet，输出是 Candidate Harness Diff 与 Mutation Report。算法不得读取密封测试题、Golden、完整 Judge 或其他算法的私有运行状态。

第一版计划包含：

- `no-evolution`：不更新 H0，用于验证控制流程。
- `prompt-reflection`：根据失败轨迹更新提示词的基础方法。
- `skill-memory`：总结经验并更新技能或记忆的基础方法。
- `full-harness`：允许修改完整做题流程的算法接口。
