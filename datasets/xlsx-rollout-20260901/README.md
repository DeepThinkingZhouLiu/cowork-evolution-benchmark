# Excel 60 题 Rollout 清单

这是 2026-09-01 Claude Code + `claude-sonnet-5` 的 60 题、每题 3 次 rollout 结果整理，不是新的 Benchmark Release。全部题包仍为 `pending_human_review`。

- `manifest.json`：运行元数据、指标口径、难度聚合和 60 条题级结果。
- `task-results.csv`：便于筛选的逐题表格；`rewards` 列按 run 顺序保存三次 reward。
- 原始汇总：`.coworkevo/runs/xlsx-60-final-complete-summary.json`，其 SHA256 已写入 `manifest.json`。

题级“Reward ≥ 0.8”采用 best-of-3（至少一次达到 0.8）；题级非零/零分按三次 reward 的均值判断。单次通过率使用 reward ≥ 0.7。总 token 和成本来自运行报告，原始题级 summary 未包含 token 字段。

重新生成：

```bash
python scripts/build_xlsx_rollout_manifest.py
```
