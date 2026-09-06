# Worktree 说明

- 用途：补齐历史题网络配置，并完成 DOCX/XLSX 共 120 道题的 Harbor Docker Oracle 回放与快照冻结。
- 分支：`lz/eval-docx-xlsx-oracle`
- 状态：`stable`
- 关键结果：DOCX 60 道与 XLSX 60 道均完成官方 Harbor 0.21 Docker Oracle 双回放；240 个 Trial 全部 Reward 1.0、0 异常，逐题两轮任务 checksum 一致。
- 下一步：把冻结快照合入主开发线，再对 DOCX/XLSX 运行三 Agent Harness Panel。
