# AutomationBench 5126 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5126()`，`example_id=5126`，任务名 `hr.safety_incident_reporting`。
- 原 Query：处理最近的工作场所安全事故，按安全程序分流与记录，使用 Google Sheets、Gmail，并向 Slack `#safety` 发布汇总。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`e42fd3eb933b27fe3d337544ea949d44ab24e5b78b48a3b462a23a94bc6193c3`。

制作时完整读取了原 Query、八条事故记录、七章安全手册、冻结时间、`#safety` 频道快照和八条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实发送邮件、更新台账、申报 OSHA、创建工伤表单或发布 Slack。

## 上游事实审计

| 员工           | 原严重度 | 手册后续状态 | OSHA / Legal                     | 工伤跟进 |
|----------------|----------|------------------|----------------------------------|----------|
| Ana Petrova    | Minor    | Documented       | 否                               | 否       |
| Ben Torres     | Moderate | Under Review     | 否                               | 是       |
| Carlos Mendoza | Serious  | Investigation    | 住院触发；Legal 立即草稿        | 是       |
| Diana Osei     | Minor    | Documented       | 否                               | 否       |
| Ethan Park     | Serious  | Investigation    | 住院触发；只能说潜在视力丧失 | 是       |
| Fiona Strand   | Minor    | Documented       | 否                               | 否       |
| George Tan     | Moderate | Under Review     | 否                               | 是       |
| Holly Ng       | Serious  | Investigation    | 住院与手指截肢触发             | 是       |

最终真值为 Minor 3 人、Moderate 2 人、Serious 3 人；3 人需立即通知 Legal 并进入 OSHA 审查，5 人需工伤跟进。源台账只给事故日期，没有事故具体时刻，因此不得虚构精确的 24 小时截止时间。公开 `#safety` 只能产生聚合草稿，不得包含员工姓名、经理邮箱、医疗细节或事故描述。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5126，并逐项断言八行事故、安全手册、冻结时间、Slack 频道和八条上游 assertions；它不是跨题通用造题器。
- Google Sheets 静态化为 `Safety_Incident_Tracker.xlsx`；Gmail 安全手册静态化为 `Safety_Manual.mbox`，保留原文与消息元数据。
- Slack 频道与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游原本要求真实发送与更新。本离线 Benchmark 只生成明确未发送、未更新、未申报、未建表和未发布的草稿与控制记录。
- 没有新增无来源的事故、人员、地址、诊断、时刻、OSHA 申报、工伤表单或外部执行事实。

## Golden 与 Judge 设计

- Golden `solve.py` 明确解析本题四个冻结输入，逐事故建立严重度、路由、状态、OSHA / Legal、工伤和隐私真值；没有随机数据或跨题通用业务生成逻辑。
- Excel 含 11 张原生工作表、11 张原生表、538 个公式单元格、4 条条件格式和 1 张可编辑图表，将源事故、10 条政策、22 条证据、三类路由、通信、Slack 与控制汇总联动。
- MBOX 恰好 8 封逐事故草稿。Tracker JSON 覆盖 8 条拟更新；Slack JSON 仅保留 1 条聚合草稿；Execution JSON 明确所有真实副作用为零。Word 为 3 页、6 张原生表格；PPT 为 3 页，每页恰好 1 张原生表格且无截图替代关键内容。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八起事故真值，并核验公式、三类路由、OSHA / Legal 优先级、工伤边界、MBOX 全部头字段与正文、三份 JSON、公开 Slack 隐私、零副作用、Word/PPT 原生结构、页数和七文件一致性。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 11 页，Word 为 3 页，PPT 为 3 页；17 页均逐页目视检查，未发现重叠、截断、越界、空白页或公式错误。
- 1,495 个本题定向负例全部被拒绝，最高得分 `0.0`，低于 `0.80` 通过线。覆盖七文件缺失/损坏、八行全字段来源、手册与 22 条证据、538 个公式单元格、三类分流、8 封邮件的全部头字段与正文、三份 JSON、OSHA 时刻缺口、工伤表单、Slack 隐私、虚假执行、Word/PPT 冲突和原生结构删除。
- 首轮完整负例暴露 22 个边界：8 个分流表工伤字段已被整行拒绝但未命中工伤 Rubric，13 个 Tracker JSON 布尔负例实际与真值相同，1 条 Word 时刻边界变异被重复文本遮蔽。Judge 与负例定位随后按题修正，完整回归为 1,495/1,495 按预期拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 第一次 Docker 预检发现 Aptos 在镜像中缺失，替代字体将 Word 从 3 页撑到 4 页，Judge 因此正确给出 `0.72`。文档字体改为宿主机和镜像均安装的 Carlito 后，两端都稳定为 3 页。
- 修复后官方 Harbor 0.21 Docker Oracle 任务 `cross-office-automationbench-5126-safety-final-v2-20260902` 与 `cross-office-automationbench-5126-safety-repeat-v2-20260902` 均为 Reward `1.0`、0 异常；trial 分别为 `automationbench-5126-safety-inci__RimdJTV`、`automationbench-5126-safety-inci__seCqyne`，任务摘要均为 `ad352b76fdb74f079717b1f7b6564f35f0f5b08f9b4875eb65ee98e5a230ff56`，七份容器产物逐字节一致。
- 正式难度暂不赋值。本题把三类分流、OSHA / Legal 优先触发、精确时刻缺口、工伤职责边界、公开频道隐私、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。
