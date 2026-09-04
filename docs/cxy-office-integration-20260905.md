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
- 四个模块均有 `environment/input_files/`；232 道任务显式设置 `network_mode = "no-network"`。

## 尚未宣称完成的事项

- PPT 60 道和 Cross-Office 60 道已有官方 Harbor 0.21 Oracle 双回放证据。
- 本次纳入的 DOCX 60 道和 XLSX 60 道已完成包级静态检查，但尚未完成统一的官方 Harbor 容器全量回放和真实 Agent Rollout。
- CXY 的部分 DOCX/XLSX 元数据仍保留来源分支的 `taxonomy_version = "v0.1"` 及旧枚举。为避免篡改来源记录，本次合并没有批量改写这些字段；正式 Release 前需按 v0.2 分类法逐题审计并生成迁移记录。
- 另有 8 道历史任务尚未显式写入 `network_mode = "no-network"`，需要单独规范化后再做全量容器回放。

因此，本次提交代表“数据已合入 `lz-dev` 的候选任务集”，不代表 240 道已经冻结为论文 Release。
