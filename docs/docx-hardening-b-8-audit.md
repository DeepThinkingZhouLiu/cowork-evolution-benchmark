# DOCX B 组 8 题复杂度升级审计

日期：2026-09-07

本记录只收录 B 组已完成“全上下文已读 -> 独立造题 -> 本地验收 -> 同一冻结快照 Harbor Oracle A/B -> commit”的 DOCX 任务。本轮不使用 VLM，difficulty 保持 `unassigned`。

| #  | 任务 | 本地 | Harbor A | Harbor B | 状态 |
|---:|------|------|----------|----------|------|
| 10 | `officebench-2-20-all-cs161` | Golden/fresh 各 5×1.0；空/损坏/图片作弊为 0；7 类定向负例拒绝；中等样例 0.970833；26 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |

## 10 officebench-2-20-all-cs161

- 改造：将只有 9 个 `Student ID, name` 段落的 v0.2 任务升级为 5 页 `CS161 Roster Control Book`。交付包含全字段源顺序名册、专业/年级分布、5 项评分对账、4 项数据质量控制、排序索引、9 个书签和 9 个内链。
- 区分度：核心名册要求逐行、逐字段与源顺序完全一致；分布、平均、极值与质量控制均从 XLSX 独立重算；导航须同时有书签和内部链接；文档需正确切换纵向 -> 横向 -> 纵向三节。
- 本地：`validate-task`、`validate-task-assets`、`compileall`、`git diff --check` 通过；Golden/fresh 各 5 次满分且逐字节一致，SHA256 `325bb9b8a1dfcb5f3763d0e4f9583a8a8671fe4d84202a3d5a68460274dd912c`。
- 负例：错学生评分 `0.35`、错源顺序 `0.35`、错专业分布 `0.68`、缺 Final 对账 `0.68`、额外学生 `0.35`、名册链接损坏 `0.72`、整页图片作弊 `0.0`；空输出和损坏文件均 `0.0`；中等样例缺 1 项质量控制，为 `0.970833`。
- 视觉：LibreOffice 另存 Golden/fresh 后仍为 `1.0`。Golden/fresh 各 5 页以及 4 份相关负例共 26 页逐页原图检查；无空白页、裁切、重叠、越界、表行断裂或字体缺失。
- Harbor A：job `docx-hardening-b8-cs161-roster-a-20260907`，trial `officebench-2-20-all-cs161__UEpjRNb`，Reward `1.0`，异常 `0`。
- Harbor B：job `docx-hardening-b8-cs161-roster-b-20260907`，trial `officebench-2-20-all-cs161__xdMFz8f`，Reward `1.0`，异常 `0`。
- 两轮一致性：`task_checksum=8d235d1bf9df88f4002f36439a91b255edf61f3dfae54ea805465c8a65ac55e4`，`task_digest=sha256:bd9e294cbd8acd50b5c504abdea6e027802fbec98f4366293c8deed568f55aa0`；产物 SHA256 均为 `325bb9b8…912c`，Judge JSON SHA256 均为 `ee321297…8833`，artifact manifest SHA256 均为 `3b2200d9…9c2c`。
- 外部证据根目录：`002-Code/agent-harnesses/runs/docx-hardening-b-8-oracle-20260907/`。
