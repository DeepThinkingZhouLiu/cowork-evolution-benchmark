# OfficeBench 3-69 成绩单 GPA 审计任务逐题适配记录

## 为什么选择这道题

上游 subtask 0 要求从 `transcripts.pdf` 抽取姓名和分数，先生成 Excel，再转成 Word。这是一条天然的“图像型 PDF -> 结构化数据 -> 可复算审计 -> Word 报告”跨 Office 链路，两份交付物的事实和结论必须一致。

结构设计目标为 L3 / Hard：Agent 需要视觉读取 23 门课程和 15 个学期，区分来源报告值与独立重算值，处理“实际交付 1 页，页脚却声明 1 of 2”的证据边界，并维持 Excel/Word 一致。按项目规则，正式难度仍保持 `unassigned`，等待冻结多 Harness Rollout 后统一校准。

## 实际读入上下文的上游材料

- `tasks/3-69/subtasks/0.json`：本题选中的完整 Query 和原 Evaluator。
- `tasks/3-69/testbed/data/transcripts.pdf`：完整打开、渲染并逐行目视核对的图像型成绩单；本题 Agent 输入为其逐字节副本。
- `utils/evaluate.py`：上游通用评测入口。
- OfficeBench 根目录 `README.md` 和 `LICENSE`，固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`。

为保留审计链，本目录保存了选中 subtask、原 Evaluator 和 Apache-2.0 许可证的精确快照。

## 逐行人工真值

| 审计对象 | 人工核对结果 |
|----------|------------------|
| 学生与签发 | Teddy Roosevelt；October 25, 2018 |
| 来源完整性 | PDF 实际 1 页；可见页脚为 `1 of 2`；来源可能不完整 |
| 来源表头 | Credits 23.00；Weighted GPA 3.78；Unweighted GPA 3.65 |
| 课程与学期 | 23 门课程；15 个学期；每门 1.00 学分 |
| 成绩分布 | 17 个 A、4 个 B、2 个 C |
| 未加权重算 | 84.00 质量分；84/23 = 3.652173913；四舍五入后 3.65 |
| 最后学期 | Drama (DE) A + World Literature (AP) B；未加权重算 3.50；来源显示 5.00 |
| 加权边界 | 可保留来源报告值 3.78，但因缺少加权政策与未交付第 2 页，不能独立重建规则 |

地址、电话、出生日期和出生地仅用于验证隐私拦截，不进入 Excel 或 Word Golden。

## 跨文件因果链

```text
transcripts.pdf
-> 23 门课程与 15 个学期的人工真值
-> Excel 课程台账与 A=4/B=3/C=2/D=1/F=0 原生公式
-> 学期级和累计 GPA 对账
-> 来源完整性与加权 GPA 证据边界
-> 与 Excel 完全一致的 Word 审计报告
```

Judge 不读取 Golden 当真值，而是内置本题逐行人工审计结果，独立检查课程顺序、成绩、标签、公式、公式重算缓存、Word 事实、证据边界、隐私和跨文件一致性。

## 对上游的必要改造

- 保留上游原 PDF 和“PDF -> Excel -> Word”的交付链，不新增学业事实。
- 上游只说“抽取姓名和分数”；本适配逐项明确为可复算课程台账、学期对账、原生图表与证据边界报告。
- 上游 Evaluator 只检查 Excel/Word 是否存在以及 Word 是否包含 `Teddy`/`Roosevelt`，无法防止漏课、错 GPA、公式静态化、编造第 2 页、隐私泄漏或跨文件冲突。
- 本题人工重写 17 条 Rubric 与任务专属 Judge，并加入 15 类定向负例。

## 许可、隐私与泄漏

来源为 Apache-2.0 的 OfficeBench 公开任务。按上游 benchmark 属性保守标记 `synthetic=true`。任何复用同一份 `transcripts.pdf` 或相同学业真值的变体都必须归入 `officebench-3-69-teddy-roosevelt-transcript` leakage group，不得跨 split。

## 验证结果

- 单题协议、输入清单与 SHA256 已通过严格校验。
- Golden 与两次空目录 fresh replay 的 Judge Reward 都是 `1.0`，两次本地产物与 Golden 逐字节一致。
- 15 个定向负例全部低于 `0.8` 通过线，最高错误得分为 `0.66`。
- XLSX 经 LibreOffice 渲染为 4 页，DOCX 渲染为 3 页，共 7 页已逐页目视检查。
- 官方 Harbor 0.21.0 Docker Oracle 独立回放两次，均为 Reward `1.0`、0 异常；两次容器产物逐字节一致。
