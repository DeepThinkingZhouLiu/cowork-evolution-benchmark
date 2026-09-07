# DOCX A 组 9 题复杂度升级进度

更新时间：2026-09-08

工作边界：仅在分支 `lz/data-docx-hardening-a-9` 与指定 worktree 中修改 A 组 9 道 DOCX 任务及直接审计证据；严禁修改 B 组第 10–17 题，不创建全局 Release，不修改根 README 或飞书 SOP，不 merge，不 push。

基线核验：

- Worktree：`REMOVED_LOCAL_PATH
- Branch：`lz/data-docx-hardening-a-9`
- 初始化基线：`a59390b0c7e2`；A/B 分组提交：`c14ab0b`
- Git status：干净（本进度文件写入前）
- Baseline pytest：`29 passed`
- Baseline public validation：`240/240 passed`
- A 组 worktree 改名后已重新读取根级 `AGENTS.md`、项目 `AGENTS.md` 和 `WORKTREE.md`；全局 Word Skill 与 Release 改由主代理在 A/B 两组完成后统一整合。

## 状态定义

- `queued`：未开始。
- `context_loaded`：已亲自读取原 instruction、全部输入、Golden、Rubric、Judge、solution、adaptation、metadata，以及实际采用的上游 Seed/Evaluator/许可。
- `authored`：已完成当题输入、Golden、Rubric、Judge、负例、来源与元数据改造。
- `local_validated`：已完成 schema/assets、Golden/fresh/empty/negative/medium、重复 Judge 与 LibreOffice 全页查看。
- `harbor_replayed`：已对同一冻结快照完成 Harbor Docker Oracle A/B 双回放，两次满分、0 异常，artifacts 齐全、digest 一致。
- `committed`：已形成可审计 commit，且未在题目内回写会使已回放快照失效的证据。

## 逐题状态

| #  | 任务 | 状态 | 当前证据 / 下一步 |
|---:|------|------|---------------------|
| 01 | `officebench-1-15-house-creak` | `committed` | 完整上下文和来源已读；v0.3 题包已造；validate/assets 通过，Golden/fresh 各5次均1.0，9类缺失/损坏/定向/中等样例符合预期，LO 另存各1.0，31页全查；Harbor A/B 各1.0、0异常、digest一致、payload稳定；题包提交 `8d9bfdb` |
| 02 | `officebench-1-15-syllabus-integrity` | `committed` | v0.3 政策整合红线工作稿完成；本地全部门禁通过、55 页全查；正式 Harbor A2/B2 各 1.0、0 异常，task checksum/digest 一致，25/25 DOCX 解压成员与 10 项 criterion 结果稳定；题包提交 `6e12e9a` |
| 03 | `officebench-1-16-random-paragraph` | `committed` | v0.3 受控文案审校工作稿完成；本地全部门禁通过、25页全查；正式 Harbor A2/B 各 1.0、0 异常，task checksum/digest 一致，22/22 DOCX 解压成员与 10 项 criterion 结果稳定；首轮 A 是未进入 Oracle/Judge 的镜像源超时，已与任务结果分开保留；题包提交 `bdffab8` |
| 04 | `officeval-003-coordination-minutes-audit` | `committed` | 旧 `officebench-1-16-test-sentence` 已用许可清楚的 OfficeVal 真实 Seed 整体替换；Golden/fresh 各5次1.0，17类缺失/损坏/定向/中等样例符合因果边，LO 另存各1.0，32页正向成品全查及6个负例焦点页实看；Harbor A/B 各1.0、0异常、checksum/digest一致、27/27 payload稳定；题包提交 `a6e8307` |
| 05 | `officebench-1-19-agreement-terms` | `committed` | Common Paper 四文件合同准备度审计已造；Golden/fresh 各5次1.0，15类负例/中等样例因果正确，LO另存各1.0，源文/正向/负例 40 页全查；Harbor A/B 各1.0、0异常、21/21 payload 稳定；题包提交 `86fcb08` |
| 06 | `officebench-1-19-paper-metadata` | `queued` | 待第 05 题闭环 |
| 07 | `officebench-1-3-rename-docx` | `queued` | 待第 06 题闭环 |
| 08 | `officebench-2-18-highest-score-statement` | `queued` | 待第 07 题闭环 |
| 09 | `officebench-2-18-shopping-list` | `queued` | 待第 08 题闭环 |

## A 组收尾门槛

- A 组 9 题汇总审计 Markdown/JSON：随逐题持续更新。
- 全库 pytest / compileall / 240 题验证 / manifest / diff-check：A 组全部题完成后执行。
- 全局 `cowork-evo-240-v0.3` Release 和 Word 造题 Skill：不在本分支创建，由主代理在 A/B 两组完成后统一整合。
