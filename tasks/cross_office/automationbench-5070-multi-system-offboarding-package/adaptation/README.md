# AutomationBench 5070 逐题改造记录

- 上游：`get_hr_offboarding_checklist_multi_app_task()` / `hr.offboarding_multi_system`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `b48d77da6fef11c2c69cd5c821f22c9dec3526d696c827fce1dd96dbc6cc6217`。
- 已完整读取原 Query、3 人 Tracker、BambooHR/Jira 初始状态和 9 条断言；`materialize_inputs.py` 只处理 5070。因 XLSX 工作表名最长 31 字符，输入表把上游 `Offboarding Processing Convention` 缩短为 `Offboarding Convention`，原文保留在精确快照。
- 人工真值：本周为 2026-03-18 至 2026-03-22，Greg 与 Ian 需进入 BambooHR/Jira/Payroll/Tracker 四系统离线草稿；Nadia 的最后工作日为 2026-04-05，必须跳过且保持 Pending。
- Golden 含 12 张 Excel 表、36 个公式、3 页 Word、3 页 PPT、6 条系统动作 JSON、1 封 MBOX 草稿和 1 份控制 JSON。Judge 独立核验日期边界、身份、PTO、Processed 约定、四系统对账、Nadia 排除、零执行与 Office 安全。
- 原 v1 的 13 类负例全部被压到 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；18 页目视检查通过；两次 Harbor Docker Oracle 回放均为 1.0、0 异常。

## 评分口径 v2（2026-09-03）

- 原严格 `rubric.json` / `tests/judge.py` / `tests/negative_cases.py` 已原样保存为 `rubric.v1.json` / `tests/judge_v1.py` / `tests/negative_cases_v1.py`；当前文件为本题专属 v2，`tests/test.sh` 默认显式选择 v2。
- v2 只在八份核心交付缺失、为空、损坏或只提交原样输入时归零；其余按三人冻结来源、12 张工作表、36 个原生公式及日期边界、Greg/Ian 四系统草稿、Nadia 排除、跨文件对账、Office 结构与离线执行边界逐项连续评分。
- 实测 Golden 和独立 fresh replay 的 v2 Reward 均为 `1.000000`，v1 Golden 为 `1.000000`。缺失、损坏、仅原样输入三类不可用样例均为 `0.000000`；14 个局部/中度变体范围为 `0.732338–0.997273`。其中四系统同时误纳 Nadia 的跨文件冲突为 `0.732338`，单一生效日错误为 `0.994667`，真实外部执行记录为 `0.944286`。
- 仓库内未找到 Claude Code、OpenCode 或 OpenHands 的本题候选 artifacts，因此本轮没有伪造三套 Harness 分数差异。

## 视觉评分适配（2026-09-03）

- 我实际重读了任务指令、三份输入、上游 5070 Query/断言、Golden 八份交付、v1/v2 Rubric 和 Judge，并查看 Golden PPT 与任务专属视觉负例的原尺寸渲染，再编写针对“范围判定 -> 两人四系统对账 -> Nadia 排除与零执行”的 `visual_rubric.json`。
- 本题是开放式离职治理交付，使用 `reference_free`：VLM 看指令、题目专属 Rubric 和候选 PPT，Golden 只用于 Rubric 制定与校准。
- 固定 Codex CLI `0.148.0` + the configured VLM provider 的 5 次完整 Golden 视觉分为 `[0.7890, 0.7974, 0.8095, 0.8550, 0.7795]`，均值 `0.806080`，总体标准差 `0.026376 < 0.05`，完整运行无重试。
- 中等视觉负例保留八份可打开交付与全部业务事实，只在第 1 页制造标题/范围窗口重叠，并把人员表向左移出页面使 Employee 列被裁切；其确定性分 `1.000000`，VLM 分 `0.615000`，混合分 `0.884500`，低于 Golden 五次均值且高于无效产物。
