# Cross-Office 逐题扩展与验收台账 v0.1

## 建设目标

Cross-Office 阶段目标 60 道已全部完成：仓库原有 10 道，后续按本文件的逐题流程新增 50 道。这里没有使用通用脚本批量制造输入、Query、Golden 或统一 Judge；每题均保留独立来源审计、真值、Rubric、Judge、负例和两次受测快照同 checksum Harbor Docker Oracle 成功回放。

## 每题必须完成的流程

1. 逐文件阅读上游 Query、全部 Agent 输入、预处理、Evaluator、Golden、数据卡和许可证；把实际读过的文件记录到 `adaptation/README.md`。
2. 冻结可公开输入并记录 SHA256。不得新增无法由上游证据支持的订单、金额、人物、结论或引用。
3. 明确跨文件因果链。至少两个应用或交付物必须共享关键事实，且上游错误需要传播到下游，不能只是把互不相关的文件放进同一目录。
4. 人工建立逐项真值台账，再设计本题专属 Golden。参考解法可以用代码落实已确认的变换，但不能用通用生成器替代逐题理解与判断。
5. 按 Harbor 目录协议提供 `task.toml`、离线 Docker 环境、`/app/input_files`、单一 `/app/output`、标准 Reward 和隐藏侧解法与评测。
6. Rubric 必须覆盖来源事实、原生 Office 结构、跨文件传播、可用性和安全边界；关键缺陷使用 hurdle cap，冲突、损坏、静态化和无来源断言使用 penalty。
7. Judge 必须从输入独立重算核心真值，并检查实际内容与结构；不得只比 Golden 哈希、文件名、关键词或页数。
8. Golden、空目录 fresh replay、空交付和定向负例全部运行；Office 文件逐页渲染目视检查。
9. 使用官方 Harbor Docker Oracle 至少独立回放两次，确认 Reward、异常数、真实 `/app/output` 和重复产物稳定。
10. 未取得有效真实Agent结果的新题，`difficulty_level` 和 `difficulty_tier` 保持 `unassigned`。当前60题已有三Harness单轮混合评分，2026-09-17已按域内Panel排名回填12/18/18/12，状态为 `pilot_single_panel`，来源见 `benchmarks/difficulty-labels-v0.5.csv`；不将该状态解释成多模型验证已经完成。

## 新增 50 道进度

| 序号 | 任务目录                                                   | 上游来源                                      | 跨文件链路                         | 负例 | Harbor 回放       | 状态              |
|-----:|------------------------------------------------------------|-----------------------------------------------|------------------------------------|-----:|-------------------|-------------------|
|    1 | `cowork-bench-financial-audit-reconciliation-package`      | Cowork-Bench `financial-audit-cross-system-reconciliation` | CSV/XLSX/PDF -> XLSX -> DOCX -> EML/ICS |   11 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    2 | `cowork-bench-conference-program-committee-package`        | Cowork-Bench `academic-conf-volga-rus-kazan-forms-teamly-excel-word-email` | Forms/Teamly -> XLSX -> DOCX/MD -> MBOX |   12 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    3 | `cowork-bench-moex-portfolio-analysis-package`             | Cowork-Bench `moex-portfolio-analysis-excel-word-email` | XLSX/CSV/PDF -> XLSX -> DOCX -> EML |   14 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    4 | `officebench-3-8-revenue-threshold-report-package`         | OfficeBench `3-8/subtask-0`                        | XLSX -> 公式分析 XLSX -> DOCX  |   13 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    5 | `officebench-3-69-transcript-gpa-audit-package`            | OfficeBench `3-69/subtask-0`                       | 图像 PDF -> 公式审计 XLSX -> DOCX |   15 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    6 | `officebench-3-91-calendar-conflict-audit-package`         | OfficeBench `3-91/subtask-0`                       | ICS + ICS -> 公式审计 XLSX -> ICS + DOCX |   19 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    7 | `officebench-3-19-20-fellowship-compliance-case-package` | OfficeBench `3-19/subtask-0` + `3-20/subtask-0` | JPG/PDF -> 公式审计 XLSX -> DOCX/EML/ICS |   21 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    8 | `officebench-3-6-team-training-coordination-package`     | OfficeBench `3-6/subtask-0`                      | XLSX -> 公式审计 XLSX -> DOCX/ICS/MBOX |   22 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    9 | `automationbench-4001-ap-invoice-control-package`        | AutomationBench `finance/4001`                   | MBOX/XLSX/Slack JSON -> 公式审计 XLSX -> DOCX/MBOX |   27 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   10 | `automationbench-1202-q1-renewal-control-package`       | AutomationBench `sales/1202`                     | MBOX/XLSX/CRM/Slack JSON -> 决策 XLSX -> DOCX/MBOX/JSON |   28 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   11 | `automationbench-5004-offboarding-control-package`      | AutomationBench `hr/5004`                        | XLSX/Policy/Slack JSON -> 决策 XLSX -> DOCX/MBOX/JSON |   32 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   12 | `automationbench-5090-confidential-access-review-package` | AutomationBench `hr/5090`                      | XLSX/MBOX/Slack/Jira -> 公式审查 XLSX -> DOCX/Jira JSON/Slack control JSON |   36 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   13 | `automationbench-5102-benefits-coverage-audit-package`  | AutomationBench `hr/5102`                        | XLSX/MBOX/Slack -> 公式覆盖审计 XLSX -> DOCX/MBOX/双 JSON |   47 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   14 | `automationbench-5104-payroll-discrepancy-control-package` | AutomationBench `hr/5104`                     | XLSX/MBOX/Slack -> 公式工资审计 XLSX -> DOCX/MBOX/双 JSON |   67 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   15 | `automationbench-5106-compliance-training-enforcement-package` | AutomationBench `hr/5106`                  | XLSX/MBOX/Slack -> 公式合规审计 XLSX -> DOCX/MBOX/双 JSON |   74 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   16 | `automationbench-5107-equipment-provisioning-package` | AutomationBench `hr/5107`                  | XLSX/MBOX/Slack -> 公式设备审计 XLSX -> DOCX/Jira JSON/MBOX/双控制 JSON |   99 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   17 | `automationbench-5108-exit-interview-privacy-control-package` | AutomationBench `hr/5108`              | XLSX/MBOX/Slack -> 公式离职审计 XLSX -> DOCX/ICS/MBOX/三控制 JSON |  137 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   18 | `automationbench-5109-referral-bonus-privacy-control-package` | AutomationBench `hr/5109`              | XLSX/MBOX/Slack -> 公式奖金审计 XLSX -> DOCX/MBOX/CSV/三控制 JSON |  142 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   19 | `automationbench-5110-org-restructure-control-package` | AutomationBench `hr/5110`              | XLSX/MBOX/Slack -> 公式改线审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON |  171 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   20 | `automationbench-5111-pto-reconciliation-privacy-package` | AutomationBench `hr/5111`           | XLSX/MBOX/Slack -> 公式 PTO 审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON |  165 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|   21 | `automationbench-5112-promotion-announcement-governance-package` | AutomationBench `hr/5112` | XLSX/MBOX/Slack -> 公式资格审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 207 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   22 | `automationbench-5113-i9-compliance-control-package` | AutomationBench `hr/5113` | XLSX/MBOX/Slack -> 公式工作日合规审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 254 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   23 | `automationbench-5115-diversity-metrics-privacy-package` | AutomationBench `hr/5115` | XLSX/MBOX/Slack -> 公式多样性审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 290 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   24 | `automationbench-5117-grievance-routing-confidentiality-package` | AutomationBench `hr/5117` | XLSX/MBOX/Slack -> 公式申诉路由审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 708 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   25 | `automationbench-5118-visa-expiration-control-package` | AutomationBench `hr/5118` | 双表 XLSX/MBOX/Slack -> 公式签证到期审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 498 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   26 | `automationbench-5119-expense-report-control-package` | AutomationBench `hr/5119` | XLSX/MBOX/Slack -> 公式费用审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 761 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   27 | `automationbench-5120-background-check-control-package` | AutomationBench `hr/5120` | XLSX/双邮件MBOX/Slack -> 公式背调审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 931 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   28 | `automationbench-5121-contractor-renewal-control-package` | AutomationBench `hr/5121` | XLSX/政策MBOX/Slack -> 公式续签审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 1064 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   29 | `automationbench-5123-job-posting-governance-package` | AutomationBench `hr/5123` | XLSX/双邮件MBOX/Slack -> 公式岗位审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 1187 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   30 | `automationbench-5126-safety-incident-control-package` | AutomationBench `hr/5126` | XLSX/安全手册MBOX/Slack -> 公式事故审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 1495 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   31 | `automationbench-5127-tuition-reimbursement-privacy-package` | AutomationBench `hr/5127` | XLSX/政策MBOX/Slack -> 公式报销审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 1577 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   32 | `automationbench-5011-salary-band-compliance-package` | AutomationBench `hr/5011` | XLSX/MBOX/Slack -> 薪资带审计 XLSX -> DOCX/PPTX/草稿/控制 JSON | 12 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   33 | `automationbench-5057-multi-app-new-hire-onboarding-package` | AutomationBench `hr/5057` | XLSX/MBOX/Slack/任务状态 -> 入职审计 XLSX -> DOCX/PPTX/MBOX/多控制 JSON | 13 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   34 | `automationbench-5070-multi-system-offboarding-package` | AutomationBench `hr/5070` | XLSX/MBOX/Slack/任务状态 -> 离职审计 XLSX -> DOCX/PPTX/MBOX/多控制 JSON | 13 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   35 | `automationbench-5101-salary-adjustment-policy-package` | AutomationBench `hr/5101` | XLSX/MBOX/Slack -> 调薪政策审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 12 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   36 | `automationbench-5103-performance-review-governance-package` | AutomationBench `hr/5103` | XLSX/MBOX/Slack -> 绩效评审审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 12 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   37 | `automationbench-5105-anniversary-recognition-governance-package` | AutomationBench `hr/5105` | XLSX/MBOX/Slack -> 司龄认可审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 12 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   38 | `automationbench-5114-termination-checklist-governance-package` | AutomationBench `hr/5114` | XLSX/MBOX/Slack -> 解聘清单审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 12 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   39 | `automationbench-5116-buddy-assignment-governance-package` | AutomationBench `hr/5116` | XLSX/MBOX/Slack -> Buddy 分配审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 13 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   40 | `automationbench-5122-survey-distribution-governance-package` | AutomationBench `hr/5122` | XLSX/MBOX/Slack -> 调研分发审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 12 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   41 | `automationbench-5124-mandatory-all-hands-governance-package` | AutomationBench `hr/5124` | XLSX/MBOX/Slack/日历状态 -> 全员会审计 XLSX -> DOCX/PPTX/MBOX/多控制 JSON | 13 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   42 | `automationbench-5125-data-migration-validation-package` | AutomationBench `hr/5125` | XLSX/MBOX/Slack -> 迁移验证 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 12 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   43 | `automationbench-5128-recognition-governance-package` | AutomationBench `hr/5128` | XLSX/MBOX/Slack -> 员工认可审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 8 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   44 | `automationbench-5129-rto-compliance-control-package` | AutomationBench `hr/5129` | XLSX/MBOX/Slack -> RTO 合规审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 8 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   45 | `automationbench-5130-headcount-confidentiality-package` | AutomationBench `hr/5130` | XLSX/MBOX/Slack -> 编制机密审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 9 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   46 | `automationbench-5131-handbook-acknowledgment-control-package` | AutomationBench `hr/5131` | XLSX/MBOX/Slack -> 手册确认审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 9 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   47 | `automationbench-5132-compensation-adjustment-governance-package` | AutomationBench `hr/5132` | XLSX/MBOX/Slack -> 薪酬调整审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 9 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   48 | `automationbench-5133-intern-onboarding-governance-package` | AutomationBench `hr/5133` | XLSX/MBOX/Slack -> 实习生入职审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 10 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   49 | `automationbench-5134-benefits-eligibility-governance-package` | AutomationBench `hr/5134` | XLSX/MBOX/Slack -> 福利资格审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 10 | 2 次，均 1.0/0 异常 | 已完成完整验收 |
|   50 | `automationbench-5135-termination-freeze-control-package` | AutomationBench `hr/5135` | XLSX/MBOX/Slack -> 终止冻结审计 XLSX -> DOCX/PPTX/MBOX/三控制 JSON | 12 | 2 次，均 1.0/0 异常 | 已完成完整验收 |

第一题的两次成功 trial 为 `cross-office-financial-audit-oracle-rerun-20260901` 和 `cross-office-financial-audit-oracle-repeat-20260901`；第二题为 `cross-office-conference-program-oracle-20260901` 和 `cross-office-conference-program-repeat-20260901`；第三题为 `cross-office-moex-portfolio-oracle-20260901` 和 `cross-office-moex-portfolio-repeat-20260901`；第四题为 `cross-office-officebench-3-8-revenue-oracle-20260901` 和 `cross-office-officebench-3-8-revenue-repeat-20260901`；第五题为 `cross-office-officebench-3-69-transcript-oracle-20260901` 和 `cross-office-officebench-3-69-transcript-repeat-20260901`；第六题为 `cross-office-officebench-3-91-calendar-oracle-fixed-20260901` 和 `cross-office-officebench-3-91-calendar-repeat-20260901`；第七题为 `cross-office-officebench-3-19-20-fellowship-oracle-20260901` 和 `cross-office-officebench-3-19-20-fellowship-repeat-20260901`；第八题为 `cross-office-officebench-3-6-training-oracle-20260901` 和 `cross-office-officebench-3-6-training-repeat-20260901`；第九题为 `cross-office-automationbench-4001-ap-oracle-20260901` 和 `cross-office-automationbench-4001-ap-repeat-20260901`；第十题为 `cross-office-automationbench-1202-renewal-oracle-20260901` 和 `cross-office-automationbench-1202-renewal-repeat-20260901`；第十一题为 `cross-office-automationbench-5004-offboarding-oracle-20260901` 和 `cross-office-automationbench-5004-offboarding-repeat-20260901`；第十二题为 `cross-office-automationbench-5090-access-review-final-20260901` 和 `cross-office-automationbench-5090-access-review-final-repeat-20260901`；第十三题为 `cross-office-automationbench-5102-benefits-audit-oracle-fixed-20260901` 和 `cross-office-automationbench-5102-benefits-audit-repeat-20260901`；第十四题为 `cross-office-automationbench-5104-payroll-control-oracle-20260901` 和 `cross-office-automationbench-5104-payroll-control-repeat-20260901`；第十五题为 `cross-office-automationbench-5106-compliance-enforcement-fixed-20260901` 和 `cross-office-automationbench-5106-compliance-enforcement-repeat-20260901`；第十六题为 `cross-office-automationbench-5107-equipment-provisioning-oracle-20260901` 和 `cross-office-automationbench-5107-equipment-provisioning-repeat-20260901`；第十七题为 `cross-office-automationbench-5108-exit-interview-privacy-oracle-20260901` 和 `cross-office-automationbench-5108-exit-interview-privacy-repeat-20260901`；第十八题为 `cross-office-automationbench-5109-referral-bonus-privacy-oracle-20260901` 和 `cross-office-automationbench-5109-referral-bonus-privacy-repeat-20260901`；第十九题为 `cross-office-automationbench-5110-org-restructure-final-20260901` 和 `cross-office-automationbench-5110-org-restructure-final-repeat-20260901`；第二十题为 `cross-office-automationbench-5111-pto-reconciliation-final-fixed-20260902` 和 `cross-office-automationbench-5111-pto-reconciliation-final-repeat-20260902`；第二十一题为 `cross-office-automationbench-5112-promotion-announcement-final-v2-20260902` 和 `cross-office-automationbench-5112-promotion-announcement-repeat-v2-20260902`；第二十二题为 `cross-office-automationbench-5113-i9-compliance-final-20260902` 和 `cross-office-automationbench-5113-i9-compliance-repeat-20260902`；第二十三题为 `cross-office-automationbench-5115-diversity-final-20260902` 和 `cross-office-automationbench-5115-diversity-repeat-20260902`；第二十四题为 `cross-office-automationbench-5117-grievance-final-20260902` 和 `cross-office-automationbench-5117-grievance-repeat-20260902`。完整来源、真值、反例与回放记录位于各题的 `adaptation/` 与 `metadata/validation.json`。

第二十五题的两次成功 trial 为 `cross-office-automationbench-5118-visa-final-20260902` 和 `cross-office-automationbench-5118-visa-repeat-20260902`。

第二十六题的两次成功 trial 为 `cross-office-automationbench-5119-expense-final-20260902` 和 `cross-office-automationbench-5119-expense-repeat-20260902`。

第二十七题的两次成功 trial 为 `cross-office-automationbench-5120-background-final-20260902` 和 `cross-office-automationbench-5120-background-repeat-20260902`。

第二十八题的两次成功 trial 为 `cross-office-automationbench-5121-contractor-final-20260902` 和 `cross-office-automationbench-5121-contractor-repeat-20260902`。

第二十九题的两次成功 trial 为 `cross-office-automationbench-5123-job-posting-final-20260902` 和 `cross-office-automationbench-5123-job-posting-repeat-20260902`。

第三十题的两次成功 trial 为 `cross-office-automationbench-5126-safety-final-v2-20260902` 和 `cross-office-automationbench-5126-safety-repeat-v2-20260902`。

第三十一题的两次成功 trial 为 `cross-office-automationbench-5127-tuition-final-20260902` 和 `cross-office-automationbench-5127-tuition-repeat-20260902`。第 32-50 题的两次受测快照同 checksum trial、产物哈希、负例和视觉检查结果位于各题 `metadata/validation.json`。

原有 10 道 Cross-Office 题也已完成 Harbor 0.21 迁移和两次受测快照同 checksum Docker Oracle 回放：6 道 Cowork-Bench、4 道 OmegaUse-OfficeVal，均为 Reward `1.0`、0 异常。迁移修复了容器缺输入、冻结 verifier 未打包、输入清单路径错误、Reward 缺失、重叠 artifacts 和慢镜像源问题。两次统一回放作业为 `cross-office-legacy-10-repeat-20260902` 和 `cross-office-legacy-10-final-repeat-20260902`，精确 trial 与 checksum 同样记录在各题 `metadata/validation.json`。

最后又对当前 60 道目录做了两轮不间断整批回放：`cross-office-60-final-a-20260902` 和 `cross-office-60-final-b-20260902`。两轮均为 60/60、均值 1.000、0 异常，每题两轮 checksum 与当前任务目录一致。每轮共产出 343 份文件；52 道逐字节一致，8 道的 OOXML 负载一致但 ZIP 时间戳或 core modified 时间不同。逐题审计清单见 [`cross-office-final-replay-20260902.json`](cross-office-final-replay-20260902.json)。
