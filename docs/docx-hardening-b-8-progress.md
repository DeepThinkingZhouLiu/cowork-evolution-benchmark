# DOCX B 组 8 题复杂度升级进度

更新日期：2026-09-07

工作边界：只在分支 `lz/data-docx-hardening-b-8` 的指定 worktree 中修改 B 组 8 道 DOCX 任务与本组审计文件；不修改 A 组、PPT、XLSX、Cross-Office、Release 或公共工具，不 merge，不 push。

状态只使用：`queued` -> `context_loaded` -> `authored` -> `local_validated` -> `harbor_replayed` -> `committed`。

| #  | 任务 | 状态 | 当前证据 / 下一步 |
|---:|------|------|---------------------|
| 10 | `officebench-2-20-all-cs161` | `committed` | 本地闭环完成；Harbor 0.21 Docker Oracle A/B 各1.0、0异常，task checksum 均为 `8d235d1b…55e4`，lock digest 均为 `sha256:bd9e294c…aa0`，两份产物与 Judge payload 均逐字节一致；任务提交 `76c262d`。 |
| 11 | `officebench-2-20-cs-major-list` | `committed` | 本地闭环完成；Harbor 0.21 Docker Oracle A/B 各1.0、0异常，task checksum 均为 `43fc33bd…2069`，lock digest 均为 `sha256:4732db31…7bd4`，两份产物与 Judge payload 均逐字节一致；任务提交 `8958d7e`。 |
| 12 | `officebench-2-20-final-perfect` | `queued` | 开始逐文件加载独立上下文。 |
| 13 | `officebench-2-21-bill-courses` | `queued` | 待 12 形成可审计 commit。 |
| 14 | `officebench-2-30-rename-student-id` | `queued` | 待 13 形成可审计 commit。 |
| 15 | `officebench-2-6-homework-ocr` | `queued` | 待 14 形成可审计 commit。 |
| 16 | `officebench-3-22-homework-folder` | `queued` | 待 15 形成可审计 commit。 |
| 17 | `officebench-3-60-cs161-folder` | `queued` | 待 16 形成可审计 commit。 |

全组最终 pytest、compileall、240/240 公开任务校验、manifest/hash 与 `git diff --check` 在 8 题单题闭环后执行。
