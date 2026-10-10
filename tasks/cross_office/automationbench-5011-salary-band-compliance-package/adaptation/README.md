# AutomationBench 5011 逐题改造记录

- 上游：`get_hr_salary_band_audit_task()` / `hr.salary_band_audit`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `b0d32626b11150c8adb95f24a4b550ae540849b9d49df95c9974c35f8ba30a20`。
- 已完整读取原 Query、5 档带宽、6 人薪酬、2 条 Flag 惯例和 8 条断言；`materialize_inputs.py` 只处理 5011。
- 人工真值：Bob 125000 相对 IC2 上界 120000 为 Over 5000；Carol 108000 相对 IC3 下界 110000 为 Under -2000；Eve 215000 相对 IC4 上界 210000 为 Over 5000；其余三人 Within。
- Golden 含 10 张 Excel 表、72 个公式、3 页 Word、3 页 PPT、1 封 MBOX 草稿和 2 份控制 JSON。Judge 独立核验带宽查找、异常、金额不舍入、零更新与 Office 安全。
- 原 v1 的 12 类负例全部被压到 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；16 页目视检查通过。

## 评分口径 v2（2026-09-03）

- 原严格口径已保存为 `rubric.v1.json` 和 `tests/judge_v1.py`；当前 `rubric.json` / `tests/judge.py` 是本题专属的 v2，`tests/test.sh` 显式调用 v2。
- v1 将任一带宽、公式、异常人员、通知或 Office 局部错误设为 hurdle，并叠加多个 penalty，导致大部分完成也被压到 0。v2 只在六份核心交付缺失、为空或损坏时归零；其余按 72 个来源单元、10 张工作表、102 个公式/语义单元、六人事实、邮件字段、两份 JSON 控制字段、Word/PPT 可编辑事实和安全边界连续计分。
- 实测 Golden 和独立 fresh replay 的 v2 Reward 均为 `1.000000`，v1 Golden 为 `1.000000`。核心文件缺失、损坏、仅原样输入三类均为 `0.000000`；12 个可解析的局部/中度错误严格位于 0 和 1 之间，范围 `0.897500–0.998039`。例如误报 Alice 为 `0.897500`，金额缩写为 `0.940000`，遗漏 Carol 邮件事实为 `0.955000`。
- 当前仓库内只能定位 oracle 的 Cross-Office 产物，没有找到 Claude Code、OpenCode 或 OpenHands 的本题候选产物，因此未伪造 Harness 差异。

## 视觉评分适配（2026-09-03）

- 我实际重读了 `instruction.md`、两份输入、上游 5011 Query/断言、Golden 六份交付、v1/v2 Rubric、Judge 和渲染后的 PPT 三页，再针对薪酬治理场景编写 `visual_rubric.json`，没有套用通用视觉模板。
- 本题是开放式成套交付，采用 `reference_free`：VLM 只看任务指令、题目专属视觉 Rubric 和候选 PPT，Golden 只用于制定与校准 Rubric。
- 固定 Codex CLI `0.148.0` + the configured VLM provider 对 Golden 真实重复 5 次，视觉分为 `[0.7780, 0.7550, 0.7150, 0.7775, 0.6825]`，均值 `0.741600`，总体标准差 `0.037400 < 0.05`，无重试。
- 中等视觉负例保留完整可打开 PPT 和全部业务事实，但故意颠倒标题/总览字号并把第 2 页表格压缩到顶部，造成明显层级、密度和空间利用问题。它的确定性分为 `1.000000`，VLM 分为 `0.500000`，混合分为 `0.850000`，严格低于 Golden 的五次均值且高于无效产物。
