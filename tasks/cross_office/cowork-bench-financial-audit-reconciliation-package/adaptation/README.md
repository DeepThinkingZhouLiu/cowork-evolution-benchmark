# Cowork-Bench 跨系统财务审计任务逐题适配记录

## 为什么选择这道题

上游任务同时要求读取 CSV、XLSX 和规则 PDF，进行逐订单核对，再交付 Excel、Word、邮件和日历状态。两份 Office 产物共享全部订单真值，邮件必须携带同一批交付物，会议邀请又必须传播最终异常结论，因此不是“多文件并列交付”，而是一条不可拆开的跨文件因果链。

结构设计目标为 L4 / Expert：三类输入、逐行对账、原生公式、双 Office 强一致、MIME 附件字节一致、ICS 外部状态草稿和输入证据边界共同构成深耦合。按照项目规则，正式难度仍保持 `unassigned`，直到 Golden、Judge、Harbor Oracle 和冻结多 Harness Rollout 全部完成。

## 实际读入上下文的上游材料

- `docs/task.md`：完整俄文 Query。
- `docs/agent_system_prompt.md`、`task_config.json`：原运行约束和工具需求。
- `initial_workspace/authoritative_ledger.csv`：14 条权威订单。
- `initial_workspace/q1_manual_tracking.xlsx`：15 条手工订单；已逐单元格读取并渲染检查。
- `initial_workspace/system_mapping_guide.pdf`：Matching Rules；已抽取文本并目视检查。
- `preprocess/main.py`：原任务只清理邮件/日历数据库，不预置答案。
- `evaluation/main.py`：原评测逐订单、汇总、邮件和日历检查。
- `groundtruth_workspace/Audit_Reconciliation_Report.xlsx`、`Audit_Report.docx`、`expected_output.json`：上游 Golden 与真值；已逐表、逐段读取并渲染检查。
- Cowork-Bench 根目录 `LICENSE`，固定 commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`。

以上可再分发快照保存在本目录。三份 Agent 输入是上游原文件的逐字节副本，`data/input_files` 与 `environment/input_files` 必须保持相同哈希。

## 人工证据台账

| 订单 | 手工证据 | 权威证据 | 结论 | 原因 |
|------|----------|----------|------|------|
| WC-001–WC-004 | 日期、金额、状态均与权威台账一致 | 同订单记录存在 | matched | none |
| WC-005 | 2024-01-15，RUB 312.25，Completed | 2024-01-15，RUB 310.00，completed | discrepancy | amount_mismatch，手工高 RUB 2.25 |
| WC-006–WC-007 | 日期、金额、状态均与权威台账一致 | 同订单记录存在 | matched | none |
| WC-008 | 2024-02-01，RUB 445.75，Completed | 无对应订单 | discrepancy | missing_in_authoritative_ledger |
| WC-009–WC-010 | 日期、金额、状态均与权威台账一致 | 同订单记录存在 | matched | none |
| WC-011 | 2024-02-15，RUB 234.00，Completed | 2024-02-20，RUB 234.00，completed | discrepancy | date_mismatch，差 5 天 |
| WC-012 | 日期、金额、状态均与权威台账一致 | 同订单记录存在 | matched | none |
| WC-013 | 2024-03-02，RUB 145.75，Completed | 2024-03-02，RUB 154.75，completed | discrepancy | amount_mismatch，手工低 RUB 9.00 |
| WC-014–WC-015 | 日期、金额、状态均与权威台账一致 | 同订单记录存在 | matched | none |

独立重算汇总：手工订单 15、matched 11、discrepancy 4、手工金额 RUB 3,263.98、已在权威台账中表示的金额 RUB 2,824.98、缺失权威记录对应手工金额 RUB 445.75、两笔金额差绝对值合计 RUB 11.25、日期异常 1、金额异常 2、缺失记录 1。

## 因果依赖图

```text
CSV + XLSX + PDF rules
-> 逐订单真值与异常原因
-> Reconciliation 公式表
-> Summary + Exception Register
-> Word 审计结论与行动
-> EML 正文 + 两个字节一致的 Office 附件
-> ICS 会议描述 + 相对附件引用
```

任何中间节点错误都会传播到后续交付物；Judge 因此同时检查来源真值、Office 原生结构、附件字节一致和四文件事实一致。

## 对上游的必要改造

- 保留全部订单、金额、日期、状态和 Matching Rules，不新增业务记录。
- 上游 Excel 只有订单号与最终状态，上游 Word 只有一段汇总和四行原因，无法验证可复算性和专业交付；本适配人工升级为三表公式工作簿和正式审计报告。
- 上游依赖真实邮件/日历数据库。为满足 Harbor 离线、可公开回放，改为 `.eml` 与 `.ics` 草稿；邮件需真正嵌入本题两个 Office 文件，ICS 使用相对 URI。
- 上游没有指定会议时间。本适配固定 2024-04-05 10:00–10:45 Asia/Shanghai，只是确定性回放约束，不作为来源业务事实。
- 不采用上游 15% 数值容差；金额匹配严格遵从规则 PDF 的零容差，Judge 对汇总金额使用货币两位小数。
- 上游只检查 Word 是否出现章节和少量关键词。本题 Judge 从输入独立重算所有订单，并检查公式、原生表格、跨文件传播、MIME 附件、ICS、渲染和无来源断言。

## 许可、隐私与泄漏

来源为 Apache-2.0 的 Cowork-Bench fixture，实体、邮箱和订单均为上游合成环境数据，因此保守标记 `synthetic=true`。本题与任何使用相同 `financial-audit-cross-system-reconciliation` 输入、订单或真值的变体必须属于同一 leakage group，不得跨 Benchmark split。

## 验证结果

- 任务协议、输入清单、三份输入 SHA256、Task Metadata、Provenance 和 Rubric Schema 全部通过。
- Golden 与空目录 fresh replay 的 Judge Reward 均为 `1.0`；11 个定向负例全部低于 `0.8` 通过阈值。
- XLSX 与 DOCX 经 LibreOffice 分别渲染为 3 页，6 页均已逐页目视检查；公式值、分页、表格和异常高亮无错误。
- 官方 Harbor 0.21.0 Docker Oracle 独立回放两次，均为 Reward `1.0`、0 异常，四个 `/app/output` 产物在两次 trial 间逐字节一致。
- 首次 Harbor 诊断回放暴露 `solve.sh` 默认写错产物目录，导致 `/app/output` 为空；修正后连续两次成功。失败与修复记录均保存在 `metadata/validation.json`，避免把基础设施问题伪装成成功结果。
