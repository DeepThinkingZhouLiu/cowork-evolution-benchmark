# DOCX 17 题复杂度升级验收记录

日期：2026-09-07

本记录只收录已完成“全上下文已读 -> 独立造题 -> 本地验收 -> 同一冻结快照 Harbor Oracle A/B”的题目。不使用 VLM 评分，正式难度保持 `unassigned`，待后续冻结 Agent Panel 校准。

## 状态表

| #  | 任务 | 本地 | Harbor A | Harbor B | 状态 |
|---:|------|------|----------|----------|------|
| 01 | `officebench-1-15-house-creak` | Golden/fresh 各 5×1.0；9 类负例/中等样例符合预期；31 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |

## 01 officebench-1-15-house-creak

### 改造摘要

- 原题只在短篇末尾追加一句固定文字；v0.3 替换为“短篇选集编辑工作稿”。
- 保护原文全部 47 个现有段落、基本格式、页面几何和页码；附录从同一授权原文综合概要、6 行风险时序与 5 角色决策分析。
- 必须交付可编辑 Word 原生对象：下一页分节、标题层级、TOC 字段、2 张表、6 个书签与内链、3 个精确锚定批注。
- Judge 从冻结输入独立重算原稿保护、故事事实、表行、批注锚点、书签/内链、原生对象和 LibreOffice 可用性，不读 Golden 哈希或解题脚本标记。

### 本地验收

- `validate-task` / `validate-task-assets` / `compileall` / `git diff --check`：通过。
- Golden 与 fresh：各重复 5 次，全部 `1.000000`；两份 DOCX 逐字节一致，SHA256 为 `700dd565a63f576ec5b6155f366f08dcbc4efadd1b731203e6fde5c974f20121`。
- 负例：缺文件 `0.0`、损坏 `0.0`、原文改字 `0.35`、错时序 `0.72`、缺角色行 `0.72`、无来源事件 `0.72`、截图式作弊 `0.0`、删批注 `0.72`、中等样例缺内链 `0.94`。
- LibreOffice 7.3.7.2 另存 Golden/fresh 后均为 `1.0`。Golden/fresh 12 页与相关负例 19 页全部逐页查看：无空白页、重叠、裁切、越界、表行断裂或字体缺失。

### Harbor Oracle A/B

- 命令环境：Harbor `0.21.0`，`agent=oracle`，`env=docker`，`n-concurrent=1`，`n-attempts=1`，`force-build=true`，容器无网。Harbor CLI 来自同项目现有 `cowork-evolution-benchmark-harbor` 环境；本地造题和校验使用指定 `cowork-evolution-benchmark` 环境。
- A job：`docx-hardening-house-creak-a-20260907`，trial `officebench-1-15-house-creak__ryobP2r`，Reward `1.0`，异常 `0`。
- B job：`docx-hardening-house-creak-b-20260907`，trial `officebench-1-15-house-creak__K9GwVf8`，Reward `1.0`，异常 `0`。
- 两次 `task_checksum` 均为 `11fd100f22073294d640120b35b82d1a814e124067a783c5fc080788aeb8b9a7`，lock task digest 均为 `sha256:769edb76df83f5e634816c31e84e0bfe03b973e622ad3b17f4c270010d98b124`。
- 两次 `/app/output` 均只含预期 DOCX，artifact manifest 均为 `status=ok`；DOCX ZIP 的所有内部 payload 逐部件完全一致，`word/document.xml` SHA256 均为 `13f8a6a99ff36c5f86935adcad02df5f4d6a89fe04e537890a99c199dbfb9c4e`。整包 SHA 因 ZIP entry 元数据时间不同而不同，不是业务 payload 漂移。
- 两次 verifier `criterion_results` 规范化 JSON SHA256 均为 `8afd2fdc3d7936c9e563f417138f1680433eef695c43a052e14e0bf85f7101f3`，10 项均满分；两份 `reward.txt` SHA256 均为 `50d3a69b930457282debea3d352ae49803a15f0843ae1699399ef4de2f4f7cdd`。
- 外部证据根目录：`002-Code/agent-harnesses/runs/docx-hardening-oracle-20260907/`。

## 待续

- 02–17 题尚未进入本记录；必须继续逐题完成，不得用通用生成器或统一模板批量构造。
- 本记录不包含三 Harness 付费基线；未获明确授权前不运行。
