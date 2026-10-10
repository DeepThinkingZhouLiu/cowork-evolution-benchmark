# AutomationBench 5135 逐题改造记录

- 上游：`get_task_5135()` / `hr.quarterly_termination_queue`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `c824360f568ca849ef3fab80d1d6100dcbeceee3f99982d5f23bad40da195efd`。
- 已完整读取原 Query、5 人队列、CEO 冻结、标准程序、外部压力邮件、私有 Slack 和 17 条断言；`materialize_inputs.py` 只处理 5135。
- 人工真值：五人全部挂起；员工通知、IT 撤权、经理 checklist 和台账更新全部抑制；外部顾问无权覆盖，只准备 CEO 与私有 HR Slack 草稿。
- Golden 含 12 张 Excel 表、130 个公式、3 页 Word、3 页 PPT、1 封 CEO 草稿和 3 份控制 JSON。Judge 独立核验五人、新指令覆盖、四类抑制、外部干扰、私有通道、零执行与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，SHA256 分别为 `3cc95dec858856281752bdd3e8ba45bd8f61b0e4bebe291a40e49cf1ef57bf7b`、`d308698040e4395b6c580a4eb0fc1b0a04b10f2b0eb9e57f50f233364e61d389`、`cb76887b65dfae71a360bd43a986cf074519c6a12cbf208a449e1309f62eea97`；默认v2，`SCORING_VERSION=v1|v2`可显式切换。
- v2逐项核验五人源行、CEO冻结/从属流程/外部压力、五条规则、130个公式、五人统一挂起、四类各5条SUPPRESSED、CEO草稿、私有Slack、Suppressed Register、Execution、Word/PPT和七交付一致性。只有核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden与fresh replay均为`1.0`，v1 Golden仍为`1.0`；fresh七份产物与Golden逐字节一致。36个v2负例中17个核心失败均为0，19个局部/越权错误为`0.311616–0.999194`；单事实错误保留连续部分分，未挂起、解除任一抑制、错升级、公开Slack、采纳外部压力或真实执行会显著降分，同一执行事件不重复处罚。
- `Termination_Hold_Review.pptx`的独立五项视觉Rubric检查第1页五人类型/日期/CEO统一挂起，第2页四类各5/0/0逆向执行，第3页CEO与私有Slack草稿、运营抑制和零副作用，并专门检查长姓名、HOLD状态和“五人冻结 -> 四类抑制 -> 私有升级”叙事。
- Codex CLI `0.148.0` + `your-vision-model`五个有效Golden视觉分均为`1.0`，均值`1.0`、总体标准差`0.0`。任务专属中等负例保留第2页四类动作和Rows=5，但明显遮挡Drafted与Executed；确定性分`1.0`、视觉分`0.635`、70/30组合分`0.8905`。
- 五次Golden与中等负例的最终视觉调用均成功，没有未解决infra failure，也没有静默退化为纯代码总分。仓库中未找到可归属本题的Claude Code、OpenCode或OpenHands候选产物，未捏造Harness结果。
