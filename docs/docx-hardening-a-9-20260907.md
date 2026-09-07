# DOCX A 组 9 题复杂度升级验收记录

日期：2026-09-07

本记录只收录 A 组 9 题中已完成“全上下文已读 -> 独立造题 -> 本地验收 -> 同一冻结快照 Harbor Oracle A/B”的题目。不使用 VLM 评分，正式难度保持 `unassigned`，待后续冻结 Agent Panel 校准。

## 状态表

| #  | 任务 | 本地 | Harbor A | Harbor B | 状态 |
|---:|------|------|----------|----------|------|
| 01 | `officebench-1-15-house-creak` | Golden/fresh 各 5×1.0；9 类负例/中等样例符合预期；31 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |
| 02 | `officebench-1-15-syllabus-integrity` | Golden/fresh 各 5×1.0；13 类负例/中等样例符合预期；55 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |

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

## 02 officebench-1-15-syllabus-integrity

### 改造摘要

- 原题只把研究诚信文本追加到课程大纲末尾；v0.3 替换为“课程政策整合红线工作稿”。
- 保护主大纲 187 个现有段落、34×4 课程表、标识、外链和页面几何；把第二份授权输入的 7 个非空政策段落放在 Turnitin 后、Special Needs 前，并使用真实 `w:ins` 修订和原生编号。
- 必须跨三种既有语境消歧：一般研究数据不当行为、Turnitin 原创性核验、校外/社区适用范围；形成 155 词 Scope Decision、7 行控制矩阵、3 条精确锚定批注、7 个书签与矩阵内链，以及独立分节页眉。
- Judge 从两份冻结输入独立重算原文保护、插入位置/顺序/修订、政策事实、矩阵行、批注锚点、导航和原生对象，不读 Golden 哈希或解题脚本标记。

### 本地验收

- `validate-task` / `validate-task-assets` / `compileall` / `git diff --check`：通过。
- Golden 与 fresh：各重复 5 次，全部 `1.000000`；整包 SHA 因 ZIP 时间戳不同，但 25/25 个解压成员逐字节一致。
- 负例：缺文件 `0.0`、损坏 `0.0`、模板改字 `0.35`、课程表损坏 `0.35`、标识替换 `0.35`、未追踪插入 `0.35`、插错章节 `0.62`、缺匿名原则 `0.62`、无依据 IRB 规则 `0.78`、匿名矩阵反向事实 `0.72`、删批注 `0.78`、截图式作弊 `0.106`、中等样例缺内链 `0.94`。
- LibreOffice 7.3.7.2 另存 Golden/fresh 后均为 `1.0`。Golden/fresh 各 9 页，相关负例 37 页，55/55 页全部逐页查看：修订位置清楚，课程表只在行间分页，附录矩阵一页完整，无空白页、重叠、裁切、越界、字体缺失或不可读内容。

### Harbor Oracle A/B

- 命令环境：Harbor `0.21.0`，`agent=oracle`，`env=docker`，`n-concurrent=1`，`n-attempts=1`，`force-build=true`；本地与容器内 Judge 都未启用 VLM。
- A job：`docx-hardening-syllabus-integrity-a2-20260907`，trial `officebench-1-15-syllabus-integr__SE4do3X`，Reward `1.0`，异常 `0`。
- B job：`docx-hardening-syllabus-integrity-b2-20260907`，trial `officebench-1-15-syllabus-integr__mmmJXus`，Reward `1.0`，异常 `0`。
- 两次 `task_checksum` 均为 `4fcefd3caeb91128b64f6782c17a51fc748a3a6cca1dea3d9abf2ac1a1881c78`，lock task digest 均为 `sha256:af6e02aa0ca18e03766cbcf22dba75c5cc9b53b166bf28c00b2a70be73e17911`。
- 两次均保留预期 output DOCX、Judge JSON、Reward 和运行日志，`/app/output` artifact manifest 为 `status=ok`；DOCX ZIP 的 25/25 个内部 payload 完全一致，`word/document.xml` SHA256 均为 `a859eb7c52b942d3eff7393f1f6b8ddc5f4d7269fb64c416e65361bb30b25140`。整包 SHA 不同只来自 ZIP entry 时间戳。
- 两次 verifier `criterion_results` 规范化 JSON SHA256 均为 `45f325959e08581d0b59e46e00efc1f5e9bdccdad0e6ee3c3bd7eb89f62cd5a2`，10 项均满分；两份 `reward.txt` SHA256 均为 `50d3a69b930457282debea3d352ae49803a15f0843ae1699399ef4de2f4f7cdd`。
- 首轮 A/B 后的提交前同步检查发现 `metadata/task.json` 漏列 task.toml 已声明的 K2；修正后重新冻结并执行上述 A2/B2。首轮结果不作为正式快照证据。
- Harbor 结束后的 LiteLLM 远端价格表刷新因 SOCKS `socksio` 缺失回退到本地备份；两次任务、Oracle、Verifier、Reward 和 artifacts 均已在该警告前成功完成，因此记录为外部非任务警告，不计为异常。
- 外部证据根目录：`002-Code/agent-harnesses/runs/docx-hardening-oracle-20260907/`。

## 待续

- 03–09 题尚未进入本记录；必须继续逐题完成，不得用通用生成器或统一模板批量构造。
- 本记录不包含三 Harness 付费基线；未获明确授权前不运行。
