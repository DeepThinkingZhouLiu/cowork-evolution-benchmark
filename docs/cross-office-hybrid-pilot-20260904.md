# Cross-Office 混合评分试跑（2026-09-04）

本轮只对三套 Harness 的已冻结 Sonnet 5 答案重评分，没有重新运行解题 Agent。视觉题使用每题独立 `visual_rubric.json`，由 Codex CLI 0.148.0 调用 `gpt-5.6-terra`；非视觉题只使用当前逐题确定性 Judge。

```text
视觉题：combined_reward = 0.7 * deterministic_reward + 0.3 * visual_reward
非视觉题：combined_reward = deterministic_reward
```

## Harness 汇总

| Harness     | 有效题 | completed | task failure | infra failure | 非零题 | 满分题 | 代码均分 | 视觉子集均分 | 混合总均分 |
|-------------|-------:|----------:|-------------:|--------------:|-------:|-------:|---------:|-------------:|-----------:|
| claude-code |     60 |        55 |            5 |             0 |     44 |      0 |    17.41% |        46.77% |      23.12% |
| opencode    |     60 |        44 |           16 |             0 |     33 |      0 |    16.76% |        49.83% |      20.34% |
| openhands   |     60 |        54 |            6 |             0 |     44 |      0 |    18.04% |        47.26% |      23.26% |

## 完整性

- 逐题结果：180 / 180。
- 正常完成：153。
- 真实任务失败：27；这些题均因受测 Agent 没有交付视觉 Rubric 指定的核心产物，不是评测基础设施故障。
- 评测基础设施失败：0；没有将 API、Codex、渲染或 JSON 解析故障记为 0 分。
- 视觉候选：108 份；成功视觉评分：81 份；其余 27 份是核心视觉产物缺失的真实任务失败。
- 表中“视觉子集均分”只对成功执行 VLM 的候选取均值；缺少核心视觉产物的 `task_failure` 没有伪造视觉分，但其混合总分按协议为 0 并计入总均分。
- 逐题明细：`benchmarks/difficulty-labels-v0.5.csv`。
- 逐题合成分已从 JSON 重算，180 份全部与 70/30 或纯代码公式一致。

## 状态边界

`task_failure` 是受测 Agent 的真实答题结果，因此保留 0 分并计入均值。`infrastructure_failure` 不产生分数且不应进入均值；本轮该项为 0。
