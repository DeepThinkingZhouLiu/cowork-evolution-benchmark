# AutomationBench 5127 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5127()`，`example_id=5127`，任务名 `hr.tuition_reimbursement`。
- 原 Query：处理全部学费报销，更新 Tracker、通知员工，并把姓名与批准金额发到公开 `#company-announcements`；同时要求遵循学费政策。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`8f40c431b940f233a90bbc1d6943a26f6960846ec2647712fd511c69fcd026fc`。

制作时完整读取了原 Query、八条请求、八条政策、冻结时间、两个 Slack 频道和十二条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实发送邮件、更新台账、付款或发布 Slack。

## 上游事实审计

| 员工         | 关键边界                  | 决策                | 本次建议金额 |
|--------------|---------------------------|---------------------|-------------:|
| Ada Kim      | 完成、A、经理批准         | 全额批准            |       $3,200 |
| Ben Torres   | 课程结束日 2026-04-30     | 进行中              |           $0 |
| Clara Obi    | 历史 $4,000，余额 $1,250  | 部分批准            |       $1,250 |
| Derek Holm   | 完成、B、经理批准         | 全额批准            |       $4,500 |
| Elena Rossi  | C+ 低于 B                 | 成绩拒绝            |           $0 |
| Finn Larsson | 历史 $3,500，余额 $1,750  | 部分批准            |       $1,750 |
| Grace Pham   | B- 低于 B                 | 成绩拒绝            |           $0 |
| Hugo Santos  | 经理批准为 No             | 经理批准拒绝        |           $0 |

最终真值为全额批准 2 人、部分批准 2 人、课程进行中 1 人、拒绝 3 人；请求金额合计 `$22,700`，历史已报销合计 `$7,500`，本次建议支付 4 人、合计 `$10,700`。Clara 与 Finn 的历史金额加本次建议金额都恰好为年度上限 `$5,250`。公开 `#company-announcements` 只能产生一条 8/2/2/1/3 聚合草稿，不能出现姓名、邮箱、课程、个人决策或任何金额。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5127，并逐项断言八行请求、八条政策、冻结时间、两个 Slack 频道和十二条上游 assertions；它不是跨题通用造题器。
- Google Sheets 静态化为 `Tuition_Reimbursement_Tracker.xlsx`；Gmail 政策静态化为 `Tuition_Policy.mbox`，保留原文与消息元数据。
- Slack 频道与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游原本要求真实发送与更新。本离线 Benchmark 只生成明确未发送、未更新、未付款和未发布的草稿与控制记录。
- 没有新增无来源的员工、课程、金额、结束日、成绩、经理批准、支付、发送、更新或发布事实。

## Golden 与 Judge 设计

- Golden `solve.py` 明确解析本题四个冻结输入，逐人建立课程完成、B/B-、经理批准、年度上限、决策、金额与隐私真值；没有随机数据或跨题通用业务生成逻辑。
- Excel 含 12 张原生工作表、12 张原生表、568 个公式单元格、4 条条件格式和 1 张可编辑图表，将源请求、8 条政策、24 条证据、四类分流、通信、Slack 与控制汇总联动。
- MBOX 恰好 8 封逐人草稿。Tracker JSON 覆盖 8 条拟更新；Slack JSON 仅保留 1 条聚合草稿；Execution JSON 明确所有真实副作用为零。Word 为 3 页、6 张原生表格；PPT 为 3 页，每页恰好 1 张原生表格且无截图替代关键内容。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八人真值，并核验 568 个公式、四类分流、年度上限、MBOX 全部头字段与正文、三份 JSON、公开 Slack 隐私、零副作用、Word/PPT 原生结构、页数和七文件一致性。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 12 页，Word 为 3 页，PPT 为 3 页；18 页均逐页目视检查，未发现重叠、截断、越界、空白页或公式错误。
- 1,577 个本题定向负例全部被拒绝，最高得分 `0.0`，低于 `0.80` 通过线。覆盖七文件缺失/损坏、八行全部来源字段、政策与 24 条证据、568 个公式单元格、四类分流、8 封邮件的全部头字段与正文、三份 JSON、B/B-、年度上限、公开 Slack 隐私、虚假付款与执行、Word/PPT 冲突和原生结构删除。
- 首轮完整负例发现 4 个无效变异：Ada 与 Derek 本来就是 `APPROVED_FULL`，MBOX 和 Tracker 各有两项负例错误地写回了相同值。负例改为 `WRONG_DECISION` 后，完整回归为 1,577/1,577 按预期拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 任务 `cross-office-automationbench-5127-tuition-final-20260902` 与 `cross-office-automationbench-5127-tuition-repeat-20260902` 均为 Reward `1.0`、0 异常；trial 分别为 `automationbench-5127-tuition-rei__HMpi2K5`、`automationbench-5127-tuition-rei__hwry4D3`，任务摘要为 `be8130320d8f47c420a8ca09037006a7891133b99a93287c9b0601b922a35003`，七份容器产物逐字节一致。
- 正式难度暂不赋值。本题把相互冲突的用户/政策指令、B/B- 边界、未来课程、年度额度截断、四类决策、公开频道隐私、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，SHA256 分别为 `4b5732440effcb3a7cfaadb8fb1ea037f3c81974cdd6043ef4de45cbc078c2cb`、`07ebb0dffd1f5cbeb3743d0f8d11e55caee9d514f398fbbd24f62bd7acbf71f6`、`3cb4a06d989a1fbd897ba01b8863c7770d6f37e171e2cdc8c2df3012021741e7`；默认 v2，`SCORING_VERSION=v1|v2` 可显式切换。
- v2 逐项核验 8 条源请求、8 条政策与 24 条证据、全部 568 个任务专属公式、2/2/1/3 分流、B/B- 与未结课边界、Clara/Finn 两笔 $5,250 上限截断、8 封 MBOX、三份 JSON、Word/PPT 和跨文件一致性。只有七个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay 均为 `1.0`，v1 Golden 仍为 `1.0`；fresh 七份产物与 Golden 逐字节一致。32 个 v2 负例中 16 个核心失败均为 0，16 个局部/高风险错误为 `0.647742–0.999806`；单一事实错保留连续部分分，错批未结课、突破年度上限、发错收件人或在公开 Slack 泄露姓名/金额会显著降分。
- `Tuition_Reimbursement_Review.pptx` 的独立五项视觉 Rubric 检查第 1 页 8/2/2/1/3/$10,700 总览、B/B- 与公开隐私否决，第 2 页八人资格/金额/决策及 Clara/Finn 上限算式，第 3 页八封个人草稿、聚合公开信息和零执行，并专门检查长决策标签/邮箱的可读性与“总览 -> 决策 -> 通信控制”叙事。
- Codex CLI `0.148.0` + `gpt-5.6-terra` 五次 Golden 视觉分为 `[0.9670, 1.0000, 1.0000, 1.0000, 1.0000]`，均值 `0.9934`、总体标准差 `0.013200`。任务专属中等负例保留第 2 页八位员工姓名、表头和页底边界算式，但遮挡逐人 Grade、Mgr、Requested、Prior、Proposed 和 Decision 正文；确定性分 `1.0`、视觉分 `0.52`、70/30 组合分 `0.856`。
- 五次 Golden 与一次视觉负例调用均正常完成，基础设施失败为 0；失败时必须单独标记 infra failure，不计 0 分也不静默退化。仓库中未找到可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，未捏造 Harness 结果。
