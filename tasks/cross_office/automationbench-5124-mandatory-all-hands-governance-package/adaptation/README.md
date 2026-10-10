# AutomationBench 5124 逐题改造记录

- 上游：`get_task_5124()` / `hr.mandatory_meeting_scheduling`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `3eb32244f6d1290f32b8786f67e19237253e94bd03ecfcc9bd220c582a44dd1a`。
- 已完整读取原 Query、8 人名单、豁免/VP+/日历政策、Slack 频道和 10 条断言；`materialize_inputs.py` 只处理 5124。
- 人工真值：Alice、Carol、Dave、Femi、Grace 必到；Bob/Hugo 因 Support 在岗豁免，Elena 因批准休假豁免；Grace 的 VP+ 覆盖被显式保留；日历修改为 0。
- Golden 含 11 张 Excel 表、112 个公式、3 页 Word、3 页 PPT、8 封 MBOX 草稿和 3 份控制 JSON。Judge 独立核验出席决策、时间、收件人、录像路由、日历越权阻断、零副作用与 Office 安全。
- 13 类负例全部 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；17 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；三份归档与改造前 HEAD 的 SHA256 分别为 `9609ecea504dea0432e5bf9074e7be4a13ad91398c302899952c3c055537cff7`、`d34a17fc4489d78562492475876626eb261ea16c8c917f9010bc2936fdc43084`、`f204c2b3c48642f9fcf6c891ba8a601883fb9fca53332a81cc37a379b1f3f8e4`。`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 不是把错误字符串个数换算成一个统一比例，而是逐项核验本题的8人源名单、政策正文、4条证据、全部112个公式、5/3出席分流、8封MBOX、三份控制JSON、Word/PPT与跨文件一致性。只有八个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与宿主机 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`；fresh 的八份产物分别与 Golden 逐字节一致。32个v2负例包含16个核心失败（均为0）和16个任务语义错误（`0.634740–0.999280`）；单字段来源、政策、公式、邮件、JSON、Word/PPT错误保留连续部分分，误邀豁免者、破坏Dave/Grace边界、虚假发送或越权改日历按风险显著降分。
- `All_Hands_Operations_Review.pptx` 使用独立五项视觉Rubric：第1页复核8/5/3/0与八人决策，第2页复核Bob/Hugo在岗豁免、Elena休假、Dave无shift note必到和Grace VP+覆盖，第3页复核5/3/1草稿、全部执行为0及日历阻断，再检查长状态可读性和“结果 -> 边界 -> 执行控制”的叙事。
- Codex CLI `0.148.0` + `your-vision-model` 的五次Golden视觉分为 `[1.000, 1.000, 0.988, 0.972, 1.000]`，均值 `0.9920`、总体标准差 `0.0110`。任务专属中等负例保留第2页五个人名和原生表格，只用白色遮罩盖住全部Decision与Reason正文，使三类豁免、Dave陷阱和Grace覆盖肉眼无法复核；确定性分仍为 `1.0`，视觉分为 `0.60`，70/30组合分为 `0.88`。
- 五次Golden和一次视觉负例调用均正常完成，基础设施失败为0；如后续调用失败，必须记录为评分基础设施失败，不得计作0分，也不得静默退化为纯确定性总分。仓库中未找到可归属本题的Claude Code、OpenCode或OpenHands候选产物，因此没有捏造Harness结果。
