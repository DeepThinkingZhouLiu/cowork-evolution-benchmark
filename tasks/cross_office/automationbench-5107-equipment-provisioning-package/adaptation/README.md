# AutomationBench 5107 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5107()`，`example_id=5107`，任务名 `hr.onboarding_equipment_provisioning`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`38a1afe50521c2e1739ca3f233930b264028984e3a2a2028dafca43e6f3ec7e6`。

本题制作时完整读取了原 Query、`initial_state.meta`、CFO 配置政策邮件全文、Google Sheets 中八名入职人员、Slack 频道、九条上游 assertions 和仓库 LICENSE。没有执行 AutomationBench 的外部服务或真实写操作。

## 上游事实审计

政策顺序为：先排除承包商/顾问，再排除备注中已发货的人员，然后按部门选型号、按地点决定显示器。逐人真值如下：

| 源行 | 人员            | 覆盖条件          | 最终状态                 | 笔记本                 | 额外显示器 |
|-----:|-----------------|-------------------|--------------------------|------------------------|-----------:|
|    2 | Ada Petrova    | 正式员工/Engineering | `PROVISION`              | Pro 16-inch M4         |          0 |
|    3 | Brian Osei     | 正式员工/Remote      | `PROVISION`              | Air 15-inch M4         |          1 |
|    4 | Carmen Delgado | Contractor        | `EXCLUDED_CONTRACTOR`    | 无                     |          0 |
|    5 | David Nguyen   | 正式员工/Engineering/Remote | `PROVISION`       | Pro 16-inch M4         |          1 |
|    6 | Elena Rossi    | 设备已由原办公室发货 | `ALREADY_PROVISIONED` | 无                     |          0 |
|    7 | Femi Adeyemi   | Independent Contractor | `EXCLUDED_CONTRACTOR` | 无                     |          0 |
|    8 | Greta Holm     | 正式员工/Support     | `PROVISION`              | Air 15-inch M4         |          0 |
|    9 | Hugo Fernandez | 正式员工/Engineering/Remote | `PROVISION`       | Pro 16-inch M4         |          1 |

最终为 5 个待配置、2 个承包商排除、1 个已配置；3 台 Pro、2 台 Air、3 台额外显示器。只生成 5 个 Jira 草稿、1 封 IT 邮件草稿和 1 条聚合频道草稿，采购、ticket 创建、邮件发送、Slack 发布和发货记录均为 0。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5107，并逐项断言八名人员、完整政策、冻结时间、频道和九条 assertions；它不是通用任务生成器。
- Google Sheets 静态化为 `New_Hire_Onboarding_Tracker.xlsx`，保留原字段、原字符串、源行号和顺序。
- Gmail 静态化为 `Provisioning_Policy_Mailbox.mbox`，保留政策全文、邮件 ID、线程 ID、日期和标签。
- Slack 与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 没有新增人员、Employee ID、部门、地点、雇佣类型、设备型号或采购事实。新增文件只用于把已审计真值变成离线、可复算、可确定检查的 Cross-Office 交付包。

## Golden 与 Judge 设计

- Golden 的 `solve.py` 明确解析本题八行人员、完整政策、已发货备注、部门和地点；没有随机数据或跨题通用业务生成逻辑。
- Excel 的来源、承包商/已发货覆盖、最终状态、型号、显示器、外设、Jira/邮件选择、队列和汇总都保留原生公式；LibreOffice 重算后再冻结包结构。
- Judge 不比较 Golden 哈希。它重新读取四个输入及 SHA256，独立构造八人真值，再检查公式依赖、缓存值、原生表格/条件格式/图表、Jira 草稿、MBOX 同行设备、Slack 聚合隐私、采购零执行、Word 同行事实和跨文件一致性。
- 关键失败使用 hurdle cap；来源/政策/设备、排除/隐私、虚假创建/发送/下单和文件/公式/追踪分别设置独立 penalty。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 6 页，Word 为 3 页；9 页均逐页目视检查，未发现重叠、截断、表格越界或公式错误。
- 99 个定向负例全部被拒绝，最高得分 `0.05`。首轮负例暴露 Word 中错误的 `ordered=true` 可被另一页同名文字掩盖，Judge 随后收紧为草稿数、四个零执行字段和最终声明必须出现在同一控制段落。
- 两次宿主机空目录 fresh replay 均为 `1.0`，六份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5107-equipment-provisioning-oracle-20260901` 和 `cross-office-automationbench-5107-equipment-provisioning-repeat-20260901` 均为 Reward `1.0`、0 异常；容器内六份产物在两个 trial 之间逐字节一致。
- 正式难度暂不赋值。本题只把多条件政策优先级、负选择、五人设备组合、三类草稿、聚合隐私和采购执行边界作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。
