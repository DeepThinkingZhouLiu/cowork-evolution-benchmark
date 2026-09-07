# agreement-terms v0.3 适配记录

状态：`local_validated_harbor_pending`；difficulty 仍为 `unassigned`。

## 替换边界

- 旧题仅从 1 页合成 PDF 复制 4 条泛化条款、2 个占位方与日期；三 Harness 收紧后均 1.0。旧 Seed 无法自然支撑冲突消解和交易就绪性分析，故整体 supersede。
- 新 Seed 是 DOCX-Redlining-Benchmark 冻结的 Common Paper 文件：CSA、DPA、Order Form 和 intake。Common Paper 素材许可为 CC BY 4.0，原文件未改动。
- 保留原“提取条款与全部参与方”能力核心，升级为 4 文件联合取证、两组优先级、20 行控制矩阵、空白变量识别、签署行动和 Word 原生导航。

## 逐文件主体已读

- 旧 instruction、PDF、Golden、rubric、judge、solver、adaptation、metadata 和三 Harness artifact/收紧结果已全文或全部 OOXML 解析。旧 PDF SHA `22802c4c...42c10`，旧 Golden SHA `c2327ef7...951c`。
- 上游 README、AGPL-3.0、Common Paper CC BY 4.0 说明、全部 `scenarios.ts` 均全文读取；commit `17382b6d8916397b4df8db6d2de7f1df6a130fe9`。
- CSA：129 段、5 表、3 节；逐段逐表行读取，包括支付、终止、删除、责任限制和 Key Terms 空白。
- DPA：82 段、3 表、2 节；逐段逐表行读取，包括 subprocessors、incident、audit、deletion 与冲突顺序。
- Order Form：4 段、2 表；intake：4 段、1 表；全部字段逐行读取。

## 真值与 Judge

- Provider/Customer/Effective Date 仅取 intake。不从其他模板推断辖区、倍数、金额或签署人。
- 15 组已明确控制与 5 组未明确控制逐项见 `evidence-ledger.md`。
- Judge 先解析 4 份冻结输入确认真值证据，再检查候选的矩阵、时限、优先级、未决项、书签内链、分节和字段；不读 Golden 哈希或 solver 标记。

## 本地验收（2026-09-08）

- `validate-task`、`validate-task-assets`、JSON 解析、单题脚本 compileall、4 份 data/environment 输入逐字节比较和 `git diff --check` 全部通过。
- Golden 与 fresh replay 各重复 5 次，10 次 reward 均为 `1.000000`；两份 DOCX 的 21/21 个解压业务部件完全一致。
- 15 个任务特有负例全部符合因果边：缺失/损坏/额外产物均为 0；当事方倒置 0.35；优先级错误均 0.55；72 小时、subprocessor 窗口或删除机制错误均 0.70；臆造责任倍数/辖区均 0.70；缺矩阵行 0.65；图片压平 0.095；只损坏横向分节为 0.986667；中等样例只缺导航为 0.92。
- LibreOffice 7.3.7.2 对 Golden/fresh 打开另存后 Judge 仍各为 1.0，两份另存件的 19/19 个解压部件一致。Golden、fresh、两份另存件各 3 页，4 份源文档共 21 页，3 个关联负例共 7 页，总计 40/40 页已逐页查看。正向文档无空白页、重叠、裁切、乱码或表行拆分；故意改回竖向的矩阵列被右侧裁掉，与 Judge 扣分原因一致。
- 本页面验收只是作者侧可用性检查，不接入 VLM，不改变 reward。Harbor Oracle A/B 待从本已冻结题包回放。
