# Benchmark 主开发线

- 用途：完整维护四模块240道题、评分代码、校准结果和Benchmark Release。
- 分支：日常开发使用 `lz-dev`；完成验证后fast-forward同步到 `main`。
- 目录：`.worktrees/origin--lz-dev--主开发线`。
- 状态：active。
- 当前产出：240道题全部回填L1-L4；统一难度索引 `benchmarks/difficulty-labels-v0.5.csv`；当前Release为 `cowork-evo-240-v0.5`，分区90/30/120。
- 校准边界：Sonnet 5三Harness单轮Panel，`pilot_single_panel`；不表示已完成独立跨模型验证。
- 下一步：从当前Release开展后续实验；题目或Judge变更后创建新Release，保留历史结果。
- 本地实验产物、凭据和缓存不提交；历史工作树保留，不再作为题目拼接入口。
