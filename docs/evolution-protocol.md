# Evolution 实验协议

## 被测对象

Evolution 算法的输入是一批允许公开的反馈，输出是对 Candidate Harness 的修改。算法不是 Judge，也不能决定自己是否晋升。

## 一个 Epoch

```text
读取当前 Harness Revision
-> 在 feedback 分区生成轨迹与评分
-> 构造 Feedback Packet
-> Evolution 算法修改 Candidate
-> 校验路径白名单与资源预算
-> 构建并运行 Candidate
-> 在 selection 分区进行配对评测
-> 按冻结 Policy 决定保留、回退或继续修订
-> 写入不可变 Evolution Ledger
```

## 公平预算

不同算法必须获得相同的：

- feedback 题目和可见字段。
- 最大 Epoch 数和 Candidate 提交次数。
- Updater Token、时间、并发和成本上限。
- Candidate 单题运行预算和失败重试规则。
- selection 查询次数和聚合反馈粒度。

算法主动提前停止时保留节省的成本，但不能把未使用预算转移到 final 评测。

## 三条赛道

- `prompt-only`：允许修改 `AGENTS.md`、提示词和少量示例。
- `skill-memory`：在上一赛道基础上允许修改技能、经验、检索索引和加载策略。
- `full-harness`：允许修改完整做题 Harness 的代码与工作流。

每条赛道由独立 Mutation Policy 声明可写和禁止路径。即使 Full Harness 赛道也不能修改 Controller、Benchmark、Judge、最终测试集、凭据和资源计量。

## 最终评测

final 分区在 Candidate 锁定前保持密封。最终报告至少包含：

- H0 在 final 上的固定基线结果。
- 最终 Harness 在相同条件下的结果。
- 成功率与平均分的配对增益。
- 所有 Epoch 的训练曲线、选择决策和成本账本。
- 最终 Harness Revision、父版本、Diff 和 Mutation Report。

禁止用 final 结果选择 Epoch、修改算法参数或再次生成 Candidate。
