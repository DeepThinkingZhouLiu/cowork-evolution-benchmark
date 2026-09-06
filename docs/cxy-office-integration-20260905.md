# CXY DOCX/XLSX 合并记录

## 来源

- 来源分支：`cxy/docx-xlsx-task-packages`
- 来源提交：`8cc5e293ec10c5772445eae24aa3bf07e1fc8155`
- 来源 PR：[#2](https://github.com/DeepThinkingZhouLiu/cowork-evolution-benchmark/pull/2)，当前仍开放，GitHub 状态为 `CONFLICTING`
- 未采用的相关分支：`cxy/harbor-migration-43`（`46a52c357e980581c3496a7fa6d101a6b5fb516e`，没有对应 PR）

## 合并结果

合并基于 `lz-dev` 的 `246e8d4` 完成，目标是把 CXY 的 DOCX/XLSX 任务数据纳入同一条开发线。

| 模块 | 合并前 | CXY 带来的变化 | 合并后 |
|------|-------:|----------------|-------:|
| DOCX | 14 | 新增 46 道 | 60 |
| XLSX | 14 | 新增 54 道，淘汰 8 个旧任务包 | 60 |
| 合计 | 28 | 净增 92 道 | 120 |

XLSX 的 8 个淘汰项是 CXY 分支明确替换的旧任务包，不是合并工具误删。`cowork-bench-moex-sector-rotation-dashboard` 的冲突文件采用 CXY 版本；其他模块没有内容冲突。

每个纳入的任务包都包含 `instruction.md`、`task.toml`、输入副本、`environment/Dockerfile`、Golden、Rubric、Judge、`solve.sh` 和 `test.sh`，并保留 `adaptation/` 与 `metadata/` 中的来源、许可、Query、哈希和改造记录。

## 已完成检查

- `cowork-evo validate-public-tasks tasks`：240/240 通过。
- 项目测试：29 个测试通过。
- `compileall`：通过。
- 四个模块目录计数：PPT 60、DOCX 60、XLSX 60、Cross-Office 60。
- 四个模块均有 `environment/input_files/`；240 道任务均显式设置 `network_mode = "no-network"`。
- DOCX 60 道与 XLSX 60 道均完成两轮官方 Harbor 0.21 Docker Oracle 回放；240 个 Trial 全部 Reward 1.0、0 异常，逐题两轮任务 checksum 一致。

## 尚未宣称完成的事项

- 四个模块各 60 道均已有官方 Harbor 0.21 Oracle 双回放证据。
- DOCX 60 道和 XLSX 60 道尚未完成真实 Agent Rollout、经验难度校准和失败样本审计。
- CXY 的部分 DOCX/XLSX 元数据仍保留来源分支的 `taxonomy_version = "v0.1"` 及旧枚举。为避免篡改来源记录，本次合并没有批量改写这些字段；正式 Release 前需按 v0.2 分类法逐题审计并生成迁移记录。
- 原有 8 道历史任务的显式 `network_mode = "no-network"` 已补齐。

因此，当前提交冻结了 240 道候选任务的 Oracle 可回放快照，但不代表已经完成真实 Agent 校准或论文最终 Release。
