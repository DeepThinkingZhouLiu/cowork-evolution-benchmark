# DOCX B 组 8 题复杂度升级审计

日期：2026-09-07

本记录只收录 B 组已完成“全上下文已读 -> 独立造题 -> 本地验收 -> 同一冻结快照 Harbor Oracle A/B -> commit”的 DOCX 任务。本轮不使用 VLM，difficulty 保持 `unassigned`。

| #  | 任务 | 本地 | Harbor A | Harbor B | 状态 |
|---:|------|------|----------|----------|------|
| 10 | `officebench-2-20-all-cs161` | Golden/fresh 各 5×1.0；空/损坏/图片作弊为 0；7 类定向负例拒绝；中等样例 0.970833；26 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `committed` |
| 11 | `officebench-2-20-cs-major-list` | Golden/fresh 各 5×1.0；14 个失败样例拒绝；中等样例 0.99；32 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `committed` |
| 12 | `officebench-2-20-final-perfect` | Golden/fresh 各 5×1.0；15 个失败样例拒绝；中等样例 0.78；22 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `committed` |
| 13 | `officebench-2-21-bill-courses` | Golden/fresh 各 5×1.0；19 个失败样例拒绝；中等样例 0.78；50 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `committed` |
| 14 | `officebench-2-30-rename-student-id` | Golden/fresh 各 5×1.0；17 个失败样例拒绝；中等样例 0.76；29 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `committed` |

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
- 任务提交：`76c262dac290aa93918410b18c76537fc8ea5172`。

## 11 officebench-2-20-cs-major-list

- 改造：将只有 3 个 `Student ID, name` 段落的 v0.2 任务升级为 5 页 `CS Major Advising Review Docket`。输出含规则排序 Review Queue、3 份逐学生 Evidence Record、逐单元格来源定位、派生计算、CS/非 CS 聚合对照、3 条原生 Word 批注、3 个书签和 3 个内链。
- 区分度：只允许披露 CS 个体，非 CS 只能以聚合形式出现；队列须按 Final 分档、Final 升序、Student ID 升序机械排序；每份记录必须把 5 项评分、工作表定位、均值、极值、极差、Midterm -> Final、趋势和 Tier 闭环。与第 10 题的全班源顺序名册目标明确不同。
- 本地：`validate-task`、`validate-task-assets`、`compileall`、JSON 与 `git diff --check` 通过；Golden/fresh 各 5 次满分且三份参考输出逐字节一致，SHA256 `e742f6a8cb777993d3a0c1864b5b03ea97b2011590df6c2118d8a38cf50a4259`；十份 Judge JSON 均为 `ffe6b122…ff7`。
- 负例：错筛选/非 CS 泄漏 `0.35`，错排序/错分档 `0.62`，错个体评分/错来源定位/缺 HW3 证据 `0.68`，错学生均值/错聚合差/缺内链/缺批注 `0.78`；空、损坏与图片压平为 `0.0`；中等样例只少 1 行冗余控制摘要，为 `0.99`。
- 视觉：LibreOffice 另存 Golden/fresh 后仍各为 `1.0`。Golden、fresh、错排序、错个体评分、非 CS 泄漏、中等样例及图片作弊共 32 页逐页原图检查；有效文档无空白页、裁切、重叠、越界、表格断裂或字体缺失。
- Harbor A：job `docx-hardening-b8-cs-major-docket-a-20260907`，trial `officebench-2-20-cs-major-list__HgzWqhh`，Reward `1.0`，异常 `0`。
- Harbor B：job `docx-hardening-b8-cs-major-docket-b-20260907`，trial `officebench-2-20-cs-major-list__ZSZNTQi`，Reward `1.0`，异常 `0`。
- 两轮一致性：`task_checksum=43fc33bd28156cba2e38562d1511a7664a6e5face754927408137a0b264e2069`，`task_digest=sha256:4732db31fe14509c2646bd04acf18dc20639fafbd5255935d0921f88828d7bd4`；产物 SHA256 均为 `e742f6a8…4259`，Judge JSON SHA256 均为 `ffe6b122…ff7`，artifact manifest SHA256 均为 `3b2200d9…9c2c`。
- 外部证据根目录：`002-Code/agent-harnesses/runs/docx-hardening-b-8-oracle-20260907/`。
- 任务提交：`8958d7e0c354f5484931a838eff35c0ee6463fb1`。

## 12 officebench-2-20-final-perfect

- 改造：将只有 2 行姓名/ID 的 v0.2 题升级为 4 页 `Perfect Final Score Verification Packet`。输出含全人口判定台账、2 份逐学生核验证书、逐单元格定位、派生计算、2 条原生 OMML 公式、阈值边界、Final 分布、Closure、2 个书签与 2 个内链。
- 区分度：必须对 9 行逐行证明 `Final = 100` 的精确等值，97 不能因“接近”而入选；每份证书要对账 5 项评分、源单元格、均值/极值/极差/排名/Final lift，最终与分布和 Closure 闭环。与第 10 题全班名册控制和第 11 题 CS 专业隐私队列的 Word 能力目标不同。
- 本地：`validate-task`、`validate-task-assets`、`compileall`、JSON 与 `git diff --check` 通过；Golden/fresh 各 5 次满分，文档 SHA256 均为 `98474d2f…e2b9`，十份 Judge JSON 均为 `62d3fa63…d4ea`。
- 负例：错阈值入选/缺台账行 `0.35`，错证书对象 `0.45`，错评分/错定位 `0.65`，错均值/缺原生公式 `0.76`，错聚合/分布/边界/内链 `0.78`；空、损坏、图片压平、多交付物均 `0.0`；中等样例只少 1 行 Closure 控制，为 `0.78`。
- 视觉：Golden/fresh 渲染页逐字节一致；Golden、fresh、97 误入选、错均值、中等样例与图片作弊共 22 页全部检查，有效页无空白、裁切、重叠、越界、表行断裂或字体缺失。LibreOffice 另存 Golden/fresh 各为 `1.0`。
- 容器网络：共享 BuildKit 长时间未出队，Debian 默认源在显式预构建时也停滞；Dockerfile 改用 Aliyun HTTPS Debian 镜像源并设有限重试/超时。用 host network 和经典 Docker 后端完成预构建，正式 Harbor A/B 仍为 `force_build=true`。
- Harbor A：job `docx-hardening-b8-perfect-final-a-20260908-pass`，trial `officebench-2-20-final-perfect__u8UEkiq`，Reward `1.0`，异常 `0`。
- Harbor B：job `docx-hardening-b8-perfect-final-b-20260908-pass`，trial `officebench-2-20-final-perfect__SQPVciL`，Reward `1.0`，异常 `0`。
- 两轮一致性：`task_checksum=adf165fa9da50a9dec16cd17c1d8b97e5f407befabc45bfc43fba0514349d30d`，`task_digest=sha256:f119259a0274c94ef14dff13728ae189c896d16fd01b992d107cd42b062b3a0b`；产物 SHA256 均为 `98474d2f…e2b9`，Judge JSON 均为 `62d3fa63…d4ea`，artifact manifest 均为 `3b2200d9…9c2c`。
- 外部证据根目录：`002-Code/agent-harnesses/runs/docx-hardening-b-8-oracle-20260907/`。
- 任务提交：`44c07698e73988c3d07184febceb797eeedaec7d`。

## 13 officebench-2-21-bill-courses

- 改造：将 15 行课程名的一页清单升级为 4 页 `Bill Course History Audit Dossier`。输出含 15 行逐字段来源台账、显式成绩点映射、五个学期及累计 GPA、三学科对账、15 行课程编号连续性图、成绩分布、十项 Closure，以及五个带稳定标签的未锁定原生 Word 富文本审核控件。
- 区分度：每门课必须同时对齐源行、学期、年份、成绩、映射点数与 `A:D` 定位；学期/学科/成绩/编号边分别独立复算再 Closure。只审计编号间距，不虚构学分、先修、毕业或学校规则，和共享来源簇的简单课程名抽取明确不同。
- 本地：`validate-task`、`validate-task-assets`、`compileall`、JSON、输入副本哈希、ownership 与 `git diff --check` 通过；Golden/fresh 各 5 次满分且逐字节一致，文档 SHA256 `5feea2b8…c60`，十份 Judge JSON SHA256 `8c977d0c…a18`。
- 负例：缺记录/错源顺序 `0.35`，错点数/定位 `0.52`，错学期/累计 `0.65`，错学科汇总/序列 `0.72`，错连续性边/定位 `0.76`，错分布或中等样例缺一项 Closure `0.78`，缺/锁定原生控件或虚构毕业事实 `0.80`；空、损坏、图片压平和多交付物均 `0.0`。
- 视觉：Golden/fresh、六份相关负例、两页图片作弊及 Golden/fresh LibreOffice 另存件共 50 页；所有变化页逐页原图检查，其余负例页与 Golden 像素字节一致。有效页无空白、裁切、重叠、越界、表行断裂或字体缺失。LibreOffice 另存保留四页、九表、五个可编辑控件值和全部业务事实，但 LO 7.3 会清空 Word 专属 `w:tag`；严格 Judge 因此给另存件 `0.80`，未掩盖互操作损失。
- Harbor A：job `docx-hardening-b8-bill-course-a-20260908`，trial `officebench-2-21-bill-courses__mfgfvqB`，Reward `1.0`，异常 `0`。
- Harbor B：job `docx-hardening-b8-bill-course-b-20260908`，trial `officebench-2-21-bill-courses__PrBUNPM`，Reward `1.0`，异常 `0`。
- 两轮一致性：`task_checksum=6832c63c30d59b4d9a62f0edb248f6f121fb74c84db0e88bf4de718fd2bfc5a0`，`task_digest=sha256:c75b1e86f59c5f9ce28c23d827902e0293c62175e18e86b80f4cfc434eb67bcf`；产物 SHA256 均为 `5feea2b8…c60`，Judge JSON 均为 `8c977d0c…a18`，artifact manifest 均为 `3b2200d9…9c2c`。
- 外部证据根目录：`002-Code/agent-harnesses/runs/docx-hardening-b-8-oracle-20260907/`。
- 任务提交：`eda1babd65b7a7690ae4cc9a783a83a1d041adc3`。

## 14 officebench-2-30-rename-student-id

- 改造：将“读取学号后复制改名”、且上游只验证文件存在的 v0.2 任务升级为三页 `Student-ID Submission Control Package`。输出包含身份派生封面、完整受保护的原作业页、13 行逐段保全台账、独立 Closure、六个自定义 Word 属性、书签/内链、页眉和自动页码。
- 区分度：必须从源段落/run 提取唯一 `A12345678`、课程、作业和学期，同时逐段、逐 run、逐格式、逐段落属性和逐编号层级保留原文；既不能只改名，也不能靠重新排版、摘要、答题或图片压平蒙混。与共享 CS161 来源簇和后续 OCR 重建簇的目标不同。
- 本地：`validate-task`、`validate-task-assets`、`compileall`、JSON、输入副本哈希、ownership 与 `git diff --check` 通过；Golden/fresh 各 5 次满分且逐字节一致，DOCX SHA256 `690e20d8…2c41`，十份 Judge JSON SHA256 `e2e02c62…0135`。
- 负例：错身份 `0.30`，受保护问题改写/缺段/合并 run `0.40`，原标题格式破坏 `0.58`，错台账指纹/缺台账行 `0.70`，中等样例缺一条冗余 Closure `0.76`，缺自定义属性/破坏导航 `0.78`；空、损坏、图片压平、错文件名、多交付物以及导致三页合同破坏的结构错误均为 `0.0`。
- 视觉：Golden/fresh、六份题目专属相关负例、两页图片作弊与两份 LibreOffice 另存件共 29 页；变化页逐页原图检查，未变化页与 Golden 像素字节一致。无空白页、裁切、重叠、越界、表格跨页、页脚冲突或字体缺失。
- 互操作：LibreOffice 7.3 另存后仍保留三页、全部正文、六表、六属性和书签/内链，但会重写源段落属性、编号定义并改变一个 run 格式签名；严格 Judge 给另存件 `0.58`，该损失未被掩盖。
- Harbor A：job `docx-hardening-b8-identity-control-a-20260908`，trial `officebench-2-30-rename-student__3ZWtzzo`，Reward `1.0`，异常 `0`。
- Harbor B：job `docx-hardening-b8-identity-control-b-20260908`，trial `officebench-2-30-rename-student__ExYy7u2`，Reward `1.0`，异常 `0`。
- 两轮一致性：`task_checksum=0a35816331fbfcb502ce1e73031b83131ab171ec4bc64ac0f640650d884e3f37`，`task_digest=sha256:fb9c5e83034fc36f6722616a7a3e183076910fa5074b2609097ca2a1d3125677`；产物 SHA256 均为 `690e20d8…2c41`，Judge JSON 均为 `e2e02c62…0135`，artifact manifest 均为 `3b2200d9…9c2c`。
- 外部证据根目录：`002-Code/agent-harnesses/runs/docx-hardening-b-8-oracle-20260907/`。
- 任务提交：`7b2ff18427d321a27059cd89a4e9c6e938220227`。
