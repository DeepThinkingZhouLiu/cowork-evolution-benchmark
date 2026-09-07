# DOCX 17 题复杂度升级进度

更新时间：2026-09-07

工作边界：仅在分支 `lz/data-docx-hardening-17` 与指定 worktree 中修改 DOCX 任务、新 Release 和直接审计证据；不修改 `tasks/ppt`、`tasks/xlsx`、`tasks/cross_office` 任务内容，不 merge，不 push。

基线核验：

- Worktree：`REMOVED_LOCAL_PATH
- Branch：`lz/data-docx-hardening-17`
- HEAD：`a59390b0c7e2`
- Git status：干净（本进度文件写入前）
- Baseline pytest：`29 passed`
- Baseline public validation：`240/240 passed`
- 旧目录改名后已重新读取根级 `AGENTS.md`、项目 `AGENTS.md`、`WORKTREE.md`、本地 PPTX 造题 Skill 与飞书 SOP 中的 PPTX/Cross-Office 两个 Skill。

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
| 02 | `officebench-1-15-syllabus-integrity` | `queued` | 待第 01 题闭环 |
| 03 | `officebench-1-16-random-paragraph` | `queued` | 待第 02 题闭环 |
| 04 | `officebench-1-16-test-sentence` | `queued` | 待第 03 题闭环 |
| 05 | `officebench-1-19-agreement-terms` | `queued` | 待第 04 题闭环 |
| 06 | `officebench-1-19-paper-metadata` | `queued` | 待第 05 题闭环 |
| 07 | `officebench-1-3-rename-docx` | `queued` | 待第 06 题闭环 |
| 08 | `officebench-2-18-highest-score-statement` | `queued` | 待第 07 题闭环 |
| 09 | `officebench-2-18-shopping-list` | `queued` | 待第 08 题闭环 |
| 10 | `officebench-2-20-all-cs161` | `queued` | 待第 09 题闭环 |
| 11 | `officebench-2-20-cs-major-list` | `queued` | 待第 10 题闭环 |
| 12 | `officebench-2-20-final-perfect` | `queued` | 待第 11 题闭环 |
| 13 | `officebench-2-21-bill-courses` | `queued` | 待第 12 题闭环 |
| 14 | `officebench-2-30-rename-student-id` | `queued` | 待第 13 题闭环 |
| 15 | `officebench-2-6-homework-ocr` | `queued` | 待第 14 题闭环 |
| 16 | `officebench-3-22-homework-folder` | `queued` | 待第 15 题闭环 |
| 17 | `officebench-3-60-cs161-folder` | `queued` | 待第 16 题闭环 |

## 全体门槛

- 新 `cowork-evo-240-v0.3` Release：未开始。
- 17 题汇总审计 Markdown/JSON：未开始。
- 全库 pytest / compileall / 240 题验证 / manifest / diff-check：未开始。
- 飞书 SOP 追加 `Skill 3：DOCX 逐题构造、作答与验收`：仅在 17 题真实闭环后执行，当前未写入。
