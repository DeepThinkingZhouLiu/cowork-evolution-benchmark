# AutomationBench 5057 逐题改造记录

- 上游：`get_hr_multi_app_new_hire_setup_task()` / `hr.multi_app_new_hire_setup`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `d7431ffb15bfc6b6c4c60b74992011d3aee221495faf4e35bf620ecdd5f0d88a`。
- 已完整读取原 Query、两位新人的七个业务字段、应用初始状态和 11 条断言；`materialize_inputs.py` 只处理 5057。
- 人工真值：Alicia 与 Tyrone 各需 BambooHR/Drive/Jira/Gmail 一条草稿；Drive 父目录为 `folder_employees`；David Okonkwo 和 Lisa Wang 仅是经理，新人档案和欢迎邮件均为0。
- Golden 含 12 张 Excel 表、32 个公式、3 页 Word、3 页 PPT、6 条系统动作 JSON、2 封 MBOX 草稿和 1 份控制 JSON。Judge 独立核验身份、字段、父目录、四系统对账、经理排除、零执行与 Office 安全。
- 原 v1 的 13 类负例全部被压到 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；18 页目视检查通过。

## 评分口径 v2（2026-09-03）

- 原严格 `rubric.json` / `tests/judge.py` 已原样保存为 `rubric.v1.json` / `tests/judge_v1.py`；当前文件为本题专属 v2，`tests/test.sh` 默认显式选择 v2。
- v1 将任一新人字段、父目录、工单、邮件或 Office 局部错误同时触发 hurdle、cap 和 penalty。v2 只在八份核心交付缺失、为空或损坏时归零；其余按两位新人的来源字段、12 张工作表、32 个公式与依赖语义、BambooHR/Drive/Jira/Gmail 逐条字段、经理排除、跨系统对账、Office 结构和离线边界连续评分。
- 实测 Golden 和独立 fresh replay 的 v2 Reward 均为 `1.000000`，v1 Golden 为 `1.000000`。缺失、损坏、仅原样输入三类不可用样例均为 `0.000000`；14 个局部/中度变体范围为 `0.837333–0.997727`。其中同时把 Lisa 误分到 BambooHR、Drive、Jira 和 Gmail 的跨系统冲突为 `0.837333`，单一父目录错误为 `0.997000`，真实外部执行记录为 `0.963846`。
- 仓库内未找到 Claude Code、OpenCode 或 OpenHands 的本题候选 artifacts，因此本轮没有伪造三套 Harness 分数差异。

## 视觉评分适配（2026-09-03）

- 我实际重读了任务指令、三份输入、上游 5057 Query/断言、Golden 八份交付、v1/v2 Rubric 和 Judge，并查看了 Golden PPT 联系表及三张原尺寸渲染，再编写针对“两人×四系统一一对账”的 `visual_rubric.json`。
- 本题是开放式入职治理交付，使用 `reference_free`：VLM 看指令、题目专属 Rubric 和候选 PPT，Golden 只用于 Rubric 制定与校准。
- 固定 Codex CLI `0.148.0` + ZCloud `gpt-5.6-terra` 的 5 次完整 Golden 视觉分为 `[0.7790, 0.8890, 0.8205, 0.7915, 0.7780]`，均值 `0.811600`，总体标准差 `0.041631 < 0.05`，完整运行无重试。
- 中等视觉负例保留八份可打开交付与全部业务事实，但把第 1–2 页标题缩小、副标题放大，并把第 2–3 页表格压缩到顶部，造成明显的层级颠倒、密度失衡和大片空白。其确定性分 `1.000000`，VLM 分 `0.780500`，混合分 `0.934150`，低于 Golden 五次均值且高于无效产物。
