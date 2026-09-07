# DOCX B 组 8 题复杂度升级审计

日期：2026-09-07

本记录只收录 B 组已完成“全上下文已读 -> 独立造题 -> 本地验收 -> 同一冻结快照 Harbor Oracle A/B -> commit”的 DOCX 任务。本轮不使用 VLM，difficulty 保持 `unassigned`。

| #  | 任务 | 本地 | Harbor A | Harbor B | 状态 |
|---:|------|------|----------|----------|------|
| 10 | `officebench-2-20-all-cs161` | Golden/fresh 各 5×1.0；空/损坏/图片作弊为 0；7 类定向负例拒绝；中等样例 0.970833；26 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `committed` |
| 11 | `officebench-2-20-cs-major-list` | Golden/fresh 各 5×1.0；14 个失败样例拒绝；中等样例 0.99；32 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `committed` |
| 12 | `officebench-2-20-final-perfect` | Golden/fresh 各 5×1.0；15 个失败样例拒绝；中等样例 0.78；22 页视觉检查 | 1.0，0 异常 | 1.0，0 异常 | `committed` |

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
