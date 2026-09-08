# OfficeBench 2-18 highest-score statement 升级记录

## 适配结论

旧题只要从 4 行成绩表取最大值，再生成一句话 DOCX；三套 Sonnet 5 Harness 均为 1.0。本题 v0.3 保留“根据成绩证据准备认可文书”这一能力核心，但把原 `score.xlsx` 重定义为提名短表，并加入同一 Apache-2.0 OfficeBench 快照中的 20 行两次期中成绩明细。这一组合自然暴露了原答案的风险：提名表只列 Rachel=100，但明细表还有 Tina=100。因此任务从“复制最大值”升级为可发布性审计、逐行对账、受控草稿和 Word 原生证据导航。

## 亲自读入的旧题与 Harness 证据

已逐项读取旧 `instruction.md`、3 个 XLSX 输入、Golden 和 solution DOCX 的全部段落/对象、`rubric.json`、`tests/judge.py`、`tests/expectations.json`、`solution/solve.py`、`solution/solve.sh`、`adaptation/README.md`、`evidence-ledger.md`、上游 Query/Evaluator、`metadata/*`与 `task.toml`。旧 Golden 只有标题和 `Rachel earned the highest score in the class: 100.`。`class_member.xlsx` 与 `shopping_list.xlsx` 与业务无关，不再暴露给解题者，字节已冻结在 `adaptation/legacy-*`。

已读取 Claude Code 2.1.251、OpenCode 1.4.14、OpenHands SDK 1.44.1 的完整轨迹、产物和收紧后 Judge 结果。Claude/OpenHands 均仅用 `openpyxl` 取最大值后写祝贺信；OpenCode 将另两个表判为无关后写两段陈述。三者都没有并列、证据范围、身份边界或发布控制推理。

## 新上游来源与许可

- 仓库：`zlwang-cs/OfficeBench`，commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`。
- 许可：Apache-2.0，完整条文保留为 `LICENSE-OfficeBench`，允许派生和分发。
- 来源1：`tasks/2-18/testbed/data/score.xlsx`，SHA256 `4a70c631...b598`；贡献 4 行提名短表及原始 Word 认可 Query。
- 来源2：`tasks/1-10/testbed/data/score.xlsx`，SHA256 `a5c8026d...d5b6e`；贡献 20 行、两次期中成绩的明细核对范围。
- Evaluator：`tasks/1-10/subtasks/0.json` 与 `1.json`，完整冻结为 `source-score-audit-*.json`；分别证实 Midterm 1 极差和 Midterm 2 下降分析是上游真实能力，新 Judge 不直接依赖其关键词。

## 真值台账与边界

- 提名短表 4 行对为 Rachel/100、Frank/67、Emma/60、Carol/97，都能在详细表找到 Name + Midterm 1 精确匹配。
- 详细表 20 条记录、15 个不同显示名；Midterm 1 最大/最小/极差为 100/60/40，Midterm 2 最大/最小为 100/63。
- Midterm 1=100 的记录有 Tina（Sheet1 行 16）和 Rachel（行 18）；Tina/100 不在提名表，所以发布状态必须是 HOLD，不能宣称 Rachel 是 sole winner。
- Alice 出现 3 次，Peter/Sam/Frank 各 2 次。源文件无学号，因此 Judge 必须以源行为实体，不允许合并重名记录。
- 未提供课程、学期、成绩等级、GPA、教师、日期、资格或学生 ID；输出不得补写。

## 适配边界

本题仍是 Word 主任务，XLSX 只是证据输入，唯一交付为可编辑 DOCX。难度来自冲突消解、逐行对账、身份边界、草稿状态一致性、原生批注、书签/内链、目录字段、分节和表格跨页控制，而不是无关文件、模糊审美或暗藏答案。不接入 VLM；LibreOffice 只用于作者侧全页验收。

## 作者侧验收观察

- `validate-task`、`validate-task-assets`、JSON 解析、compileall、两份 data/environment 输入字节比对和 `git diff --check` 均通过。
- Golden 与 fresh replay 各 5 次 Judge 均为 1.0，两者 DOCX SHA256 同为 `082351cf6a23b1d363401fa851683c4d79aad69d03316d2ad51ea70f178480a1`，22/22 个解压成员字节相同。
- LibreOffice 打开/另存后 Golden 和 fresh 仍各得 1.0；两个另存件 20/20 个解压成员相同。
- 全页实际查看 20 页：2 页 XLSX 源证据、Golden/fresh 各 3 页、5 个重点负例共 12 页。修复过首轮第 2 页只留一条内链的近空白页问题；最终正向稿为 3 页，首页完整容纳发布状态、对账和两份草稿，第 2 页横向完整容纳 20 行台账和重名控制，第 3 页是发布控制；无空白页、裁切、重叠、行拆分、缺字或不可读内容。横向改竖向负例的台账右侧明显被裁，压平负例只剩黑色位图，视觉失败与 Judge 因果一致。
- 视觉验收不进入 reward，不存在 `visual_rubric`；正式难度仍为 `unassigned`，等待后续冻结 Panel 标定。
