# DOCX B 组 8 题复杂度升级进度

更新日期：2026-09-08

工作边界：只在分支 `lz/data-docx-hardening-b-8` 的指定 worktree 中修改 B 组 8 道 DOCX 任务与本组审计文件；不修改 A 组、PPT、XLSX、Cross-Office、Release 或公共工具，不 merge，不 push。

状态只使用：`queued` -> `context_loaded` -> `authored` -> `local_validated` -> `harbor_replayed` -> `committed`。

| #  | 任务 | 状态 | 当前证据 / 下一步 |
|---:|------|------|---------------------|
| 10 | `officebench-2-20-all-cs161` | `committed` | 本地闭环完成；Harbor 0.21 Docker Oracle A/B 各1.0、0异常，task checksum 均为 `8d235d1b…55e4`，lock digest 均为 `sha256:bd9e294c…aa0`，两份产物与 Judge payload 均逐字节一致；任务提交 `76c262d`。 |
| 11 | `officebench-2-20-cs-major-list` | `committed` | 本地闭环完成；Harbor 0.21 Docker Oracle A/B 各1.0、0异常，task checksum 均为 `43fc33bd…2069`，lock digest 均为 `sha256:4732db31…7bd4`，两份产物与 Judge payload 均逐字节一致；任务提交 `8958d7e`。 |
| 12 | `officebench-2-20-final-perfect` | `committed` | 4 页 Final=100 例外核验包已完成本地闭环；Harbor 0.21 Docker Oracle A/B 各1.0、0异常，task checksum 均为 `adf165fa…d30d`，lock digest 均为 `sha256:f119259a…3a0b`，产物/Judge/manifest 逐字节一致；任务提交 `44c0769`。 |
| 13 | `officebench-2-21-bill-courses` | `committed` | 四页课程历史审计档案已本地闭环；Harbor 0.21 Docker Oracle A/B 各 1.0、0 异常，task checksum 均为 `6832c63c…c5a0`，lock digest 均为 `sha256:c75b1e86…7bcf`，产物/Judge/manifest 逐字节一致；任务提交 `eda1bab`。 |
| 14 | `officebench-2-30-rename-student-id` | `committed` | 三页身份派生/受保护内容保全包已本地闭环；Harbor 0.21 Docker Oracle A/B 各 1.0、0 异常，task checksum 均为 `0a358163…3f37`，digest 均为 `sha256:fb9c5e83…5677`，产物/Judge/manifest 逐字节一致；任务提交 `7b2ff18`。 |
| 15 | `officebench-2-6-homework-ocr` | `committed` | 三页可编辑 OCR 重建、流水线矩阵与证据/冲突审计包已本地闭环；Harbor 0.21 Docker Oracle A/B 各 1.0、0 异常，task checksum 均为 `dca0475e…d95`，digest 均为 `sha256:e9bc2730…86d1`，产物/Judge/manifest 逐字节一致；任务提交 `5edea5f`。 |
| 16 | `officebench-3-22-homework-folder` | `committed` | 四页受控 OCR 主副本 + 四页 QA 工作副本已闭环；Harbor 0.21 Docker Oracle A/B 各 1.0、0 异常，task checksum 均为 `e4ca090e…7770`，digest 均为 `sha256:cc5f80a2…3355`，两份角色化 DOCX、Judge 与 manifest 逐字节一致；任务提交 `cf0b0f5`。 |
| 17 | `officebench-3-60-cs161-folder` | `committed` | 11 页长名册核验登记簿已本地闭环；Harbor 0.21 Docker Oracle A/B 各 1.0、0 异常，task checksum 均为 `cd77458e…a26e`，digest 均为 `sha256:7fafdcfd…6ab9`，DOCX/Judge/manifest 逐字节一致；任务提交 `a7deb62`。 |

全组终验已完成：`PYTHONPATH=src` 下 29 个 pytest 全过；`src`、`tests`、`tools` 与 B 组 8 题的 solution/tests 全部 compileall 通过；240/240 公开任务校验通过；8 题 input manifest、provenance、运行时输入副本、Golden/solution 哈希一致；ownership、修改边界、`git diff --check` 与工作树清洁性通过。若盲扫整个 `tasks/`，会命中未改动 PPT 任务中原样冻结的上游 evaluator f-string 语法错误，该文件不属于 B 组且没有被修改。
