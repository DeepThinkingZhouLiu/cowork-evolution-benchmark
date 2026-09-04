# AutomationBench 5105 逐题改造记录

- 上游：`get_task_5105()` / `hr.anniversary_recognition`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `a808c8817e9659869707d6f5dadc41924667cb4825e6bf3035185f66c352f0b7`。
- 已完整读取原 Query、10 人目录、历史认可格式、VP 高管 CC 政策、Frank 入职日更正、Slack 上下文和 9 条断言；`materialize_inputs.py` 只处理 5105。
- 人工真值：Angela 3年、Ben 5年、Frank 2年、Grace 2年、Jake 4年；Celine/Derek 非4月，Eva/Henry 日期无效，Irene 不在职；禁止高管 CC。
- Golden 含 12 张 Excel 表、200 个公式、3 页 Word、3 页 PPT、1 封 all-staff MBOX 草稿和 3 份控制 JSON。Judge 独立核验更正、月份、年资、排除、格式、收件人、CC 压制、零执行与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。

## v2 连续评分与视觉校准（2026-09-04）

- 原严格口径保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；默认 v2，`SCORING_VERSION=v1|v2` 可显式选择。
- v2 保留十人目录、Frank 后到日期更正、五名4月庆祝者、五名按原因排除者、高管CC禁令、all-staff草稿和零执行的任务专属检查。只有七份核心交付缺失、为空、损坏、不可解析或只交原样输入硬归零，其余事实、公式、结构和跨文件错误连续计分。
- all-staff收件人、高管Cc、压制记录、邮件发送/副作用和Word虚假声明按9个沟通治理原子项连续扣分；排除者误入邮件、遗漏Frank及更正日期/年资/应用错误按7个选择原子项连续扣分。Golden/fresh v2 与 v1 Golden 均为 `1.0`；14个 v2 负例中3个致命负例为`0`，其余11个为`0.610000–0.930000`。
- PPTX 采用本题独有5项视觉Rubric，覆盖五人庆祝表、Frank更正与三类排除、高管CC压制/离线边界、长姓名和状态可读性及整套认可治理风格。Codex CLI `0.148.0` + `gpt-5.6-terra` 五次 Golden 视觉分为 `0.9340、0.9375、0.9505、0.8957、0.9405`，均值 `0.93164`、总体标准差 `0.018793`。
- 中等视觉负例保持可打开且确定性分 `1.0`，但将第2页副标题叠到标题、证据表整体右移并裁掉决策列；视觉分 `0.7200`，低于 Golden 均值且高于无效产物，70/30 合成分 `0.9160`。一次中断目录保留但未计入统计；未发现三套 Harness 候选产物。
