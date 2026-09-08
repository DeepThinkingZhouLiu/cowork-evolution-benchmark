# OfficeBench 2-18 shopping-list 升级记录

## 适配结论

旧题仅需把 3 行 `Name + amount` 从 Excel 搬到 Word，三套 Sonnet 5 Harness 均为 1.0。v0.3 保留“根据购物数据交付可编辑 Word 文档”的核心，但将同一 OfficeBench 购物 Seed 家族中的数量变更、删除、新增和定价要求冻结为四份可见控制记录，组成一次真实的采购交接与变更审计。

## 亲自读入的旧题与 Harness 证据

已逐项读取旧 `instruction.md`、data/environment 输入、Golden 与 solution DOCX 的段落/表格/OOXML 部件、`rubric.json`、`tests/judge.py`、`tests/expectations.json`、`solution/solve.py`、`solve.sh`、adaptation 文档、上游 Query/Evaluator、`metadata/*`和 `task.toml`。旧 Golden 只有一张 4×2 表，内容是 Name/Amount 表头与 apple/cabbage/onion 各3行，没有推理、差异、追溯或 Word 原生审阅对象。旧 data/environment/Golden/solution/source Query 已原样移入 `adaptation/legacy/`。

已读取 Claude Code 2.1.251、OpenCode 1.4.14 和 OpenHands SDK 1.44.1 的完整轨迹、实际产物及收紧后 Judge。Claude/OpenCode 都是读 3 行后直接创建表格；OpenHands 经过 pandas 缺失后改用 openpyxl，亦只复制 3 行并验证文件存在。三份产物都完全缺少变更序列、删除留痕、定价计算、来源定位和 Word 原生审阅对象。

## 新上游来源与贡献

- 仓库：`zlwang-cs/OfficeBench`，commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`；Apache-2.0，完整许可保留为 `LICENSE-OfficeBench`。
- `tasks/2-18/testbed/data/shopping_list.xlsx`，SHA256 `e6468e54...821a`：贡献 3 行基线清单与原始 Word 转写能力。
- `tasks/1-5/subtasks/1.json`，SHA256 `f08f630b...4062`：贡献 apple 减半的变更语义与 evaluator 中的目标量。
- `tasks/1-7/subtasks/5.json`，SHA256 `4a31a7b4...69a6`：贡献 onion 删除语义和原值证据。
- `tasks/1-4/subtasks/1.json`，SHA256 `7f9f65ab...c1d3`：贡献 yogurt 新增及数量 evaluator。
- `tasks/1-6/subtasks/5.json`，SHA256 `8930b9d5...7ec1`：贡献 apple 例外价与“others”通用价规则。

## 真值与适配边界

基线为 apple=2、cabbage=2、onion=2。按 T01 -> T02 -> T03 各应用一次后，当前表为 apple=1、cabbage=2、yogurt=2，顺序为保留的基线项后追加新项；onion 只作为 removed 记录留在审计表。定价规则明文以美元/磅表述，因此本题仅为成本计算将 amount 视为磅；“others”价同样适用于批准新增的 yogurt。这是 instruction 明示的 composite 适配规则，不改写上游数值。

本题仍是 Word 主任务：XLSX/JSON 仅作证据，唯一交付为可编辑 DOCX。难度来自变更次序、基线与当前态分离、逐行对账、成本重算、删除留痕、原生批注、书签/内链/目录与分节表格，而不是模糊审美、无关输入或暗藏答案。Judge 每次从 5 个冻结输入独立重算，不读 Golden 哈希或 solver 标记；本题不接入 VLM。

## 作者侧验收

- `validate-task`、`validate-task-assets`、compileall、JSON 解析、5 份 data/environment 输入逐字节比对和 `git diff --check` 均通过。
- Golden 与独立 fresh replay 各运行 5 次 Judge，10 次均为 1.0；所有文件 SHA256 均为 `2f5ccd358dd211a46213f9b2fe0b797e17ee035ca09fc5950e082804a5649f78`，24/24 个 OOXML 解压成员字节一致。
- 19 个故障/中等样例结果：缺文件、损坏、多交付、错文件名均 0；T01 重复应用、onion 回流、yogurt 成本错均 0.48；T02 台账缺行与票据乱序均 0.62；onion 状态错与 yogurt 对账缺行均 0.68；总成本错 0.72；定价来源缺失 0.76；批注未锚定 0.78；页码字段缺失和控制表非横向均 0.79；虚构供应商/税 0.75；截图压平 0.125；仅缺导航的中等样例 0.955。
- LibreOffice 打开/另存后 Golden 和 fresh 仍各得 1.0；两个另存件的 22/22 个解压成员字节一致。
- 实际查看 32 页：1 页 XLSX 基线、Golden/fresh 各 3 页，以及 9 个重点负例的 25 页。正向稿为竖向首页 -> 横向三张控制表 -> 竖向计算/边界页，无空白页、裁切、重叠、拆行、缺字或不可读内容。非横向负例的当前表、变更台账和对账表右侧明显被裁；压平负例只剩一块黑色位图，视觉因果与 Judge 一致。
- 本轮不设 `visual_rubric`，上述查看只是作者侧可用性验收，不进入 reward。Harbor Oracle A/B 将在本题冻结快照后运行，结果只写入题外 A 组审计，不再改变被测快照。

正式难度保持 `unassigned`，只能由后续冻结 Panel 标定。
