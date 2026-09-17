# PPT 任务难度试校准（2026-09-04）

## 结论

本轮对 60 道 PPT 任务重新评分并洗牌。每道题使用 Claude Code 2.1.251、OpenCode 1.4.14 和 OpenHands SDK 1.44.1 各一份已有 Sonnet 5 答案，代码 Judge 保留原分，视觉部分由 Codex CLI 0.148.0 调用 `gpt-5.6-terra` 按每题独立 `visual_rubric.json` 评分。视觉题的最终分为：

```text
combined_reward = 0.7 * deterministic_reward + 0.3 * visual_reward
```

三套 Harness 的任务级最终分取平均，得到 `panel_mean`；再按 `panel_mean` 从高到低排序，依照预先约定的 12 / 18 / 18 / 12 配比划为 Easy / Medium / Hard / Expert。边界没有同分；若未来出现同分，先由 PPT 结构难度规则审计，仍无法区分时再按稳定任务编号排序。

## 分档结果

| 难度   | 题数 | Claude Code | OpenCode | OpenHands | Panel 均分 | 经验难度 |
|--------|-----:|------------:|---------:|----------:|-----------:|---------:|
| Easy   |   12 |       96.96 |    97.29 |     97.98 |      97.41 |     2.59 |
| Medium |   18 |       35.51 |    37.27 |     40.61 |      37.80 |    62.20 |
| Hard   |   18 |       18.02 |    16.53 |     16.23 |      16.93 |    83.07 |
| Expert |   12 |        9.19 |     6.37 |      7.11 |       7.56 |    92.44 |

表中分数均为百分制；经验难度为 `100 - Panel 均分`。三个 Harness 的档位均值都满足 Easy > Medium > Hard > Expert。

## 运行完整性

- 评分结果：180 / 180。
- 正常完成：179。
- 真实任务失败：1；OpenCode 在 `officeval-075-template-style-transfer` 未交付要求的 PPTX，因此记 0。
- 评测基础设施失败：0；没有把 API、Codex、渲染或解析故障记为 0。
- 本地完整评分 artifacts：`.coworkevo/empirical-calibration/ppt-hybrid-20260904/panel/`。
- 逐题排名与分数：`benchmarks/ppt-difficulty-pilot-20260904.csv`。

## 状态边界

这是一轮试校准：三套Harness使用同一Sonnet 5主干，每题各有一份有效答案。2026-09-17已将上述CSV的既有档位回填到60题的 `task.toml` 和 `metadata/task.json`，统一清单为 `benchmarks/difficulty-labels-v0.5.csv`；分数与题目归属不变，校准状态明确记为 `pilot_single_panel`。标签固定用于当前版本分析，不代表已完成独立跨模型或重复运行验证；后续验证应加入更多模型家族和重复运行，基础设施失败不进入难度统计。
