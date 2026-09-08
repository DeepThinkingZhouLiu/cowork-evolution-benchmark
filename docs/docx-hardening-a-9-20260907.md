# DOCX A 组 9 题复杂度升级验收记录

日期：2026-09-07

本记录只收录 A 组 9 题中已完成“全上下文已读 -> 独立造题 -> 本地验收 -> 同一冻结快照 Harbor Oracle A/B”的题目。不使用 VLM 评分，正式难度保持 `unassigned`，待后续冻结 Agent Panel 校准。

## 状态表

| #  | 任务 | 本地 | Harbor A | Harbor B | 状态 |
|---:|------|------|----------|----------|------|
| 01 | `officebench-1-15-house-creak` | Golden/fresh 各 5×1.0；9 类负例/中等样例符合预期；31 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |
| 02 | `officebench-1-15-syllabus-integrity` | Golden/fresh 各 5×1.0；13 类负例/中等样例符合预期；55 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |
| 03 | `officebench-1-16-random-paragraph` | Golden/fresh 各 5×1.0；12 类缺失/损坏/定向/中等样例符合预期；25 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |
| 04 | `officeval-003-coordination-minutes-audit` | Golden/fresh 各 5×1.0；任务特有负例闭环；38 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |
| 05 | `officebench-1-19-agreement-terms` | Golden/fresh 各 5×1.0；15 类负例/中等样例；40 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |
| 06 | `officebench-1-19-paper-metadata` | Golden/fresh 各 5×1.0；16 类负例/中等样例；23 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |
| 07 | `officeval-088-product-release-dossier` | Golden/fresh 各 5×1.0；18 类负例/中等样例；28 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |
| 08 | `officebench-2-18-highest-score-statement` | Golden/fresh 各 5×1.0；17 类负例/中等样例；20 页视检查 | 1.0，0 异常 | 1.0，0 异常 | `harbor_replayed` |

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

## 03 officebench-1-16-random-paragraph

### 改造摘要

- 原题只把 5 句随机段落复制成一个 Word 段落；v0.3 整体替换为“咖啡馆文案受控审校工作稿”。
- 三份输入分工清晰：原 TXT 是唯一业务事实源，Word 模板提供受保护原文、横竖分节、台账和汇总结构，批准票据仅提供 5 条事实保守改写与 3 条精确批注。
- 交付物必须保留受控源，生成 5 对真实 `w:del`/`w:ins`、3 个原生批注、5 个书签内链、6×6 句子台账、103 -> 68 -> -35 词数汇总和可追溯证据说明。
- Judge 从 TXT 和票据 DOCX 独立重算真值，逐修订、逐批注、逐表格字段、逐书签和原生结构评分，不读 Golden 哈希或解题器标记。

### 本地验收

- `validate-task` / `validate-task-assets` / `compileall` / JSON 解析 / 3 份输入字节一致 / `git diff --check`：通过。
- Golden 和 fresh 各重复 5 次，全部 `1.000000`；整包时间戳不同，但 22/22 个解压成员一致。
- 负例：缺失 `0.0`、损坏 `0.0`、额外文件 `0.0`、受控源改字 `0.35`、接受修订 `0.58`、缺 S5 修订 `0.58`、S3 错批准文案 `0.58`、错词数/差值 `0.70`、删批注 `0.72`、无据价格/营业时间 `0.76`、图片压平 `0.1125`、破坏横向分节 `0.984444`；中等样例缺导航 `0.93`。
- LibreOffice 7.3.7.2 另存 Golden/fresh 后均为 `1.0`。Golden/fresh、另存和相关负例共 25 页全部逐页看图：正确文档无空白页、重叠、裁切、表行断裂或字体缺失；破坏分节和压平负例的视觉失败与 Judge 因果一致。

### Harbor Oracle A/B

- 正式 A2 job：`docx-hardening-random-paragraph-a2-20260908`，trial `officebench-1-16-random-paragrap__KJWaoH5`，Reward `1.0`，异常 `0`。
- 正式 B job：`docx-hardening-random-paragraph-b-20260908`，trial `officebench-1-16-random-paragrap__wK2Gmw7`，Reward `1.0`，异常 `0`。
- 两次 `task_checksum` 均为 `77aeecacf7e59ad3b7c14e2c035b1b36950a5a1a9ed164a4fbf076e6545d5c85`，lock task digest 均为 `sha256:b0ce4d8c6545fa15e370252e39d038df3c84684ac6921ed8069b6a3a5b817148`。
- 两次 `/app/output` 均只含 `cafe_profile_editorial_review.docx`，artifact manifest 均为 `status=ok`；DOCX 的 22/22 个内部 payload 逐部件一致，`word/document.xml` SHA256 均为 `6ac449b6cf5aa0aeceb09ce9a4238eca943c4292c2a36a752b7af2cf5c8ca7ef`。
- 两次 verifier `criterion_results` 规范化 JSON SHA256 均为 `280cabb5ff1c8dd73767862f9622f0b751e945dfdd457a3f80b76f3e39c0b462`，10 项全满分；`reward.txt` SHA256 均为 `50d3a69b930457282debea3d352ae49803a15f0843ae1699399ef4de2f4f7cdd`。
- 首轮 `docx-hardening-random-paragraph-a-20260907` 在旧 Debian HTTP 镜像下构建 30 分钟后发生 `EnvironmentStartTimeoutError`，0 个有效 trial，未进入 Oracle/Judge。该 Job 保留为基础设施证据；题目改用同步 HTTPS 镜像并保留 apt 重试后，上述 A2/B 为正式冻结快照。
- 两次结束后的 LiteLLM 远程价格表刷新因 SOCKS `socksio` 缺失回退到本地备份；任务、Verifier、Reward 和 artifacts 均已先成功，因此记为外部非任务警告。
- 题包提交：`bdffab8`；外部证据根目录：`002-Code/agent-harnesses/runs/docx-hardening-oracle-20260907/`。

## 04 officeval-003-coordination-minutes-audit

### 迁移与改造摘要

- 迁移映射：`officebench-1-16-test-sentence` -> `officeval-003-coordination-minutes-audit`。旧题只要求写入一句固定测试句，输入 TXT 与用户请求无关，因此不在旧 Seed 后堆要求，而是用 Apache-2.0 的 OmegaUse-OfficeVal `officeval_003` 真实 Word Seed 整体 supersede。
- 新题保留“把会议内容整理为细黑横线表”核心，并从同一冻结会议记录递归取证：拆除表中表和59个空段/行，复原161条语义记录，核对5条作业偏差和8项决议。
- 交付物是一份可编辑 DOCX：包含未接受的“若干 -> 5”真实修订、精确锚定批注、13个书签/内链、横向审计附录、独立页眉和 PAGE/NUMPAGES 页脚。
- Judge 从冻结源 DOCX 独立重算，逐字段、逐行、逐修订/批注、逐表格、逐书签评分；不读 Golden 哈希，不使用 VLM。

### 本地验收

- `validate-task` / `validate-task-assets` / compileall / JSON 解析 / 2份输入 data/environment 逐字节比对 / `git diff --check`：全部通过。
- Golden 和 fresh 各重复5次，全部 `1.000000`；两份 DOCX 的 27/27 个解压业务成员完全一致，整包 SHA 差异仅来自 ZIP 时间戳。
- 负例：缺失 `0.0`、损坏 `0.0`、额外文件 `0.0`、空转写行 `0.65`、偏差数改6 `0.55`、批注锚点删除 `0.55`、接受修订 `0.55`、表决顺序错 `0.75`、偏差责任人错配 `0.70`、虚构责任人/时限 `0.75`、首行红色粗线 `0.65`、图片压平 `0.12`、破坏横向分节 `0.9925`、嵌套表回归 `0.55`、地点破坏 `0.35`、漏转写行 `0.55`；中等样例只缺13个导航链接，得分 `0.93`。
- LibreOffice 7.3.7.2 另存 Golden/fresh 后均为 `1.0`，另存件 33/33 个解压业务成员一致。Golden、fresh、两份另存件各8页，32/32页逐页实看；正向成品无裁切、重叠、空白页、乱码或表行拆分。另聚焦查看6个版面负例页，红粗线、图片压平、嵌套表和竖向附录裁列均与 Judge 因果一致。

### Harbor Oracle A/B

- A job：`docx-hardening-coordination-minutes-a-20260908`，trial `officeval-003-coordination-minut__MtJ4wEG`，Reward `1.0`，异常 `0`。
- B job：`docx-hardening-coordination-minutes-b-20260908`，trial `officeval-003-coordination-minut__MN6awMH`，Reward `1.0`，异常 `0`。
- 两次 `task_checksum` 均为 `d2844bb638d6013ce33c6b63d14287512ef93a57b45b7f0b32fd322313fa3046`，lock task digest 均为 `sha256:5ad3fd509115539a54b9d878680a948781197b895ec36a87e954c2da9da8d92e`。
- 两次 `/app/output` artifact manifest 都为 `status=ok`；27/27个 DOCX 解压业务成员完全一致，`word/document.xml` SHA256 都为 `3e98179e0b35375da173ee09f35abc2c1a611edb2cb6741b4b92b3e858d7ee19`。整包 SHA 分别为 `fa5c2d79...83c7` 和 `e0515f5b...a144`，差异仅来自 ZIP 时间戳。
- 两次 verifier `criterion_results` 规范化 JSON SHA256 均为 `fd2f3184628c7ccc0c3d5ba76c9667d6e3a3cd06bf2b43831320ed11fe11d01f`，10项均满分；题包提交 `a6e8307`。
- Harbor 对任务约定与显式 `--artifact /app/output` 报告了同路径重复采集警告；manifest 中两个记录均为 `ok`，实际产物只保留一份，不是任务异常。

## 05 officebench-1-19-agreement-terms

### 改造摘要

- 原题只从 1 页合成 PDF 复制 4 条泛化条款、2 个占位方与日期；v0.3 用 CC BY 4.0 的 Common Paper CSA、DPA、Order Form 和 intake 四文件整体 supersede，保留“条款与当事方提取”能力核心。
- 新题需跨文档消解两组优先级，区分 15 组有证据控制与 5 组未明确变量，并抑制从拟稿示例臆造责任倍数、金额、辖区或签署权。
- 交付物含 20 行条款/准备度矩阵、5 项签署行动、20 个书签内链、TOC 字段、横向分节和 PAGE/NUMPAGES 页脚字段；Judge 从四份冻结输入独立校验真值。

### 本地验收

- `validate-task` / `validate-task-assets` / compileall / JSON / 4 份 data/environment 逐字节比较 / `git diff --check`：全部通过。
- Golden 和 fresh 各 5 次均 1.0，21/21 个解压业务部件一致；15 类缺失/损坏/定向/中等样例符合预期。
- LibreOffice 7.3.7.2 另存 Golden/fresh 后仍各 1.0。19/19 个另存 payload 部件一致。Golden、fresh、另存件、四份源文和三个相关负例合计 40/40 页逐页实看；正向成品无空白页、裁切、表行拆分、重叠或乱码。

### Harbor Oracle A/B

- A job `docx-hardening-agreement-terms-a-20260908`，trial `officebench-1-19-agreement-terms__UBjtfE3`；B job `docx-hardening-agreement-terms-b-20260908`，trial `officebench-1-19-agreement-terms__mYGFfkf`。两次 reward 均 1.0，异常均 0。
- 两次 `task_checksum` 均为 `a1c3f921d10e1ac2b4c0e7f39e1152024b274255056d8db61ec8eb120960b83c`，lock task digest 均为 `sha256:761e4f696293ab0c7fcb79b9d528ee7058cfbcca4e0f19c8a312803aa00bde67`。
- 两份 DOCX 的 21/21 个解压部件完全一致，`word/document.xml` SHA256 均为 `dce0f3a2dd9ecae97b50be5775e37efb124acb8dfde56f599f01850b28185a9a`；criterion 结果规范化 SHA256 均为 `e12014576fb3d9a128c68529fd229219cde4faa8d4e7b3fb98e535777e7eacbf`。
- 两次 artifact manifest 中 `/app/output` 都为 `ok`，只含预期 DOCX。同路径重复采集警告来自 task 约定与显式 `--artifact` 重合，不是任务异常。题包提交：`86fcb08`。

## 06 officebench-1-19-paper-metadata

### 改造摘要与本地验收

- 保留 OfficeBench 原始一页 OS-Copilot PDF 及标题/作者提取核心，自然升级为研究入库与证据审计：解析 8 名作者的符号单位映射、等贡献边界、摘要主张、Figure 1 三场景，并显式限制一页证据不支持的 DOI、会议、baseline 或统计细节。
- 成品含 6 张原生表格、3 分节、12 个书签内链、3 个原生批注、TOC、核心属性和页码字段。Judge 从 PDF 文本与 `pdfinfo` 独立取证，不读 Golden。
- 两项 validator、JSON、compileall、输入字节比较与 diff-check 通过；Golden/fresh 各5次均1.0，24/24 payload 一致；16 类负例/中等样例符合预期；LO 另存各1.0且 22/22 payload 一致。源文、正向、另存与 4 个负例共 23/23 页实看，正向成品无空白、裁切、表行拆分、重叠或乱码。

### Harbor Oracle A/B

- A job `docx-hardening-paper-metadata-a-20260908`，trial `officebench-1-19-paper-metadata__GNwbPZU`；B job `docx-hardening-paper-metadata-b-20260908`，trial `officebench-1-19-paper-metadata__ohodzDi`。两次均 1.0，异常均 0。
- `task_checksum` 均为 `fba96a8711dde5b8f5bfcc263db15575bddfd312158f3729b59b73267fa58ce2`，lock digest 均为 `sha256:b404b934aae01aa54feabf75db3fb66034882fff88943d41ff6ce48a78602603`。
- 24/24 个 DOCX payload 部件完全一致，`word/document.xml` SHA256 均为 `724d1e38684303a41ef8a982eb6fbd5deb66071ae3b3967f085da8ac4c6daf8e`；criterion 结果 SHA256 均为 `53f3afb8cd82fe811a1f7cc5ee7e08a90535ae49830921c1a3f39eb2e1b217f0`。两次 artifact manifest 对 `/app/output` 均为 `ok`。题包提交 `19db0be`。

## 07 officeval-088-product-release-dossier

### 迁移与改造摘要

- 迁移映射：`officebench-1-3-rename-docx` -> `officeval-088-product-release-dossier`。旧题的唯一输入只含 `PDF test file 2`，三套 Harness 都一次复制后满分，因此用 Apache-2.0 的 OfficeVal 088 七文件真实 Seed 整体替换。
- 新题保留“以指定 Word 文件名发布”的能力核心，但发布动作落在五份产品说明取证、指导价排名、前三模板填充、产品图裁剪、五产品证据矩阵、价格差额复算和证据边界控制之后。
- 交付物只有一份可编辑 `产品分析.docx`；XLSX 仅作为字段语义输入。成品包含 4 张原生表格、横向分节、3 个精确批注、8 个书签、6 条内链、TOC、PAGE/NUMPAGES 和可访问图片替代文本。Judge 每次从五份来源独立重算，不读 Golden。

### 本地验收

- `validate-task` / `validate-task-assets` / compileall / JSON / 7 份 data/environment 输入逐字节比较 / `git diff --check`：全部通过。
- Golden/fresh 各 5 次均 `1.0`，整包及 17/17 个解压成员完全一致；18 类缺失、损坏、额外文件、错名、模板破坏、排序/事实/证据/图片/批注/分节/页码域/虚构/压平负例和中等样例均符合预期。
- LibreOffice 7.3.7.2 另存 Golden/fresh 后仍各 `1.0`，18/18 个另存 payload 一致。七份源输入 13 页、Golden/fresh/另存件 8 页和四个关键负例 7 页，共 28/28 页逐页查看；正向文档稳定为 2 页，无空白、裁切、重叠、表行拆分、缺字或不可读内容。

### Harbor Oracle A/B

- A job `docx-hardening-product-release-a-20260908`，trial `officeval-088-product-release-do__gha3sHr`；B job `docx-hardening-product-release-b-20260908`，trial `officeval-088-product-release-do__NkRPSEr`。两次 reward 均 `1.0`，异常均为 0。
- 两次 `task_checksum` 均为 `f5912cd3011e23be173553d17e279ce49a75faa59dbca44783239f3bd7e70106`，lock digest 均为 `sha256:ba1f4ac8194a834dea6ea384d7605436816a234cea8bae1cc8160a731e102caf`。
- 两份 DOCX 的 17/17 个解压成员完全一致，`word/document.xml` SHA256 均为 `fcc7e795545fdc29aeecaf210ce8e17b2599cc1f90f683d5665678270e37d65c`；10 项 criterion 规范化 SHA256 均为 `cf0a32c50177bb09956e4610bf0dae7c2f337bd0bc6ea5493d9e08fcfdb457af`。
- 两次 artifact manifest 对 `/app/output` 都为 `ok`，只含预期 DOCX；重复路径提醒仍是采集配置去重而非任务异常。题包提交 `8e7c25e`。

## 08 officebench-2-18-highest-score-statement

### 改造摘要与本地验收

- 原题只从 4 行表取 Rachel=100 并写一句话；v0.3 将这份表定义为提名短表，再加入同许可 OfficeBench 1-10 的 20 行两次期中成绩明细，构成可发布性审计。
- 新证据显示 Tina 与 Rachel 在 Midterm 1 并列 100，而 Tina 被提名表遗漏；成品必须 HOLD，不得宣称单一赢家。Alice/Peter/Sam/Frank 重名但无学号，因此 20 条记录必须逐行保留，不可合并。
- 唯一交付为 3 页可编辑 DOCX：4 行提名对账、20 行带符号差值台账、2 份受控草稿、2 处原生批注、29 个书签、6 条内链、TOC/PAGE/NUMPAGES 字段和横向台账分节。Judge 从两份 XLSX 每次独立重算。
- 两项 validator、JSON、compileall、输入字节比较和 diff-check 通过；Golden/fresh 各 5 次 1.0，22/22 payload 一致，DOCX SHA256 均为 `082351cf6a23b1d363401fa851683c4d79aad69d03316d2ad51ea70f178480a1`。
- 17 类负例/中等样例覆盖缺失、损坏、额外文件、错名、虚假发布、漏并列者、单一赢家、错极差、缺对账行、合并重名行、错差值、删批注、去页码字段、改竖向、虚构 ID/GPA 和图片压平，为 0–0.79；只缺导航的中等样例为 0.97。
- LibreOffice 另存 Golden/fresh 后均为 1.0，20/20 payload 一致。来源、正向和 5 个重点负例共 20 页全部查看；首轮发现并修复了“Rachel 内链单独溢出成近空白页”，最终正向无空白、裁切、重叠、拆行或缺字。

### Harbor Oracle A/B

- A job `docx-hardening-recognition-review-a-20260908`，trial `officebench-2-18-highest-score-s__wXqeLZC`；B job `docx-hardening-recognition-review-b-20260908`，trial `officebench-2-18-highest-score-s__8m5fks9`。两次 reward 均为 1.0，异常均为 0。
- 两次 `task_checksum` 均为 `098d2a37450c9fa04ada8410f228edd642f4580213ddd1e3bc059bd6c70c4603`，lock digest 均为 `sha256:39fa0fee7a5da68f2364edc011be3935c244b893c38720772f06a1c21bc0435a`。
- 两份 DOCX 整包及 22/22 个解压成员完全一致，`word/document.xml` SHA256 均为 `328228d8b486ead50d7a5a0d56adcdd82f85b8a5a15d1df5971b3efcaac0680c`；10 项 criterion 规范化 SHA256 均为 `98dbf5e2ba15f12fa3125e92045e54e1062ccb5cb6ec095729f02e164e9bc8c4`。
- 两次 `/app/output` manifest 均为 `ok`，只含预期 DOCX；Judge JSON、reward、lock 和日志齐全。重复路径提醒是 task 约定与显式 `--artifact` 重合后去重，不是任务异常。题包提交 `34d1c45`。

## 待续

- 09 题尚未进入本记录；必须继续逐题完成，不得用通用生成器或统一模板批量构造。
- 本记录不包含三 Harness 付费基线；未获明确授权前不运行。
