# PPTArena Case 71 逐题适配记录

## 适配契约

- 输入是 4 页演示文稿，其中第 2–4 页各有 1 张原生可编辑表格。
- 真实意图是识别 6 处采用半角减号、Unicode 减号或括号表示的负数，将至少数字设为红色粗体。
- 主交付物是 `ConditionalTableFormatting_Negatives_GroundTruthA.pptx`。
- 所有表格文字必须原样保留，正数和非目标文字不得改样式，三张表格保持原生可编辑。
- 候选设计目标为 L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout。

## 上游来源与文件审计

- 来源：PPTArena Case 71: Conditional Table Formatting (Negative Numbers)。
- GitHub commit：`6f26a63ec1014b98c44990c7b58133b832070be2`；Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`；许可为 MIT。
- 本题直接使用上游原生 PPTX 和 Query，不是本项目编造的 Mock 输入。PPTArena 数据卡将数据标记为公开 MIT 数据，本题业务数字为合成示例，已如实写入 provenance。

| 文件                                                                   | 字节数 | SHA256                                                             | 实际检查                                                     |
|--------------------------------------------------------------------------|-------:|--------------------------------------------------------------------|--------------------------------------------------------------|
| `data/input_files/ConditionalTableFormatting_Negatives_TestA.pptx`      |  36847 | `0b64628a3f1169dc6dcfca2ea0acca80e1cc6aeba7b9aa9bc456e8f36b9a305f` | 4 页逐页渲染；3 张表格、56 个单元格、全部 run 与 6 处负数审计 |
| 上游 `ConditionalTableFormatting_Negatives_GroundTruthA.pptx`   |  37157 | `653f6bd3641d08e8f2694e57e5e7ce6d7713bf88f700ec2508fb5671cf433f64` | 负数样式真值正确；同时擅自改封面并把 `(1.2%)` 改成 `-1.2%`     |
| `solution/Golden/ConditionalTableFormatting_Negatives_GroundTruthA.pptx` |  31904 | `a19d9fda1134653cb7fd2a12cf94ca2d719ccae04c26d431a0a8d112e5640228` | 从输入逐单元格构建；只对 6 处负数设红色粗体，文字不变      |
| `adaptation/upstream-case.json`                                         |   1866 | `48bdd3f8fd683716c8f4b128150e20ff5686c929be85e67e9d3c10710f238913` | Query、style target、路径和 enhancement note 逐字段比对       |
| `tests/fixtures/golden-slide2.png`                                      |  38272 | `b5bb9feb96391b3042e2fb08e333165f015e1b4a882d78059868f6c0e7e2df07` | 只用于第 2 页压平负例，Judge 不将其当作答案                      |

## Golden 与 Judge 设计

- `build_golden.py` 锁定输入 SHA256，显式定位六个单元格，保留原 run 的语言与字体属性，仅将目标字符设为 `#FF0000` 和粗体。
- Judge 不读取 Golden。它从冻结输入独立检查 6 处负数的每个数字是否红色粗体，并允许负号、货币符号和标点可选同步格式化。
- 非目标单元格与目标外字符逐字样式对比；表格非文本 OOXML 经规范化后对比，同时检查行列、几何、备注、非表格对象和整页压平。
- 独立 LibreOffice profile 渲染输入与候选，遮罩六个目标单元格后检查页面保持，并从目标单元格计算可见红色像素。

## 实际验证结果

| 样例                                      | 实际 reward | 结果 | 关键原因                                      |
|-------------------------------------------|------------:|------|-----------------------------------------------|
| 本地 Golden                              |    1.000000 | 通过 | 9 条判据全部通过                             |
| 只将数字设红色粗体的合法变体              |    1.000000 | 通过 | 证明 Judge 不依赖 Golden 的 run 分段          |
| fresh replay                              |    1.000000 | 通过 | 空目录回放产物哈希与 Golden 一致               |
| 未修改 / 漏一处 / 只红不粗 / 只粗不红 |    0.000000 | 拒绝 | 六处目标的完整性、颜色或字重门槛失败       |
| 误改正数 / 全表染色 / 改数据              |    0.000000 | 拒绝 | 非目标逐字样式或数据保护失败                 |
| 移表格 / 整页截图 / 夹带答案标记          |    0.000000 | 拒绝 | 几何、原生可编辑性或交付真实性失败            |
| 额外输出 `debug.txt`                   |    0.960000 | 部分 | 主交付正确，但输出卫生正分失败                  |

同一 fresh replay 产物判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena071-oracle-negative-tables-20260829`，trial 为 `pptarena-071-negative-number-tab__gQB4RrM`，0 个异常；容器产物 SHA256 与本地 Golden 一致。

## 视觉检查与剩余风险

- 输入、上游 Ground Truth、本地 Golden 和合法变体均已逐页渲染查看；6 处红色粗体负数清晰可见，表格无裁切或位移。
- `visual_qa.py` 对输入、Golden 和合法变体均只报告封面原有的两条文本框稀疏警告，三者完全相同，不是本题新增问题。
- 正式 Release 前仍建议用 Microsoft PowerPoint 打开/保存抽检；L2 只是制题目标，正式难度必须经真实 Agent Rollout。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Case 与元数据、`build_golden.py` / `build_fixtures.py` / `solve.py`、冻结输入和本地 Golden；核对了三张表、六个负数的精确单元格和三种负数表示法，以及现有的漏改、误改、移动、压平、夹带标记等逐题负例。
- 冻结输入、本地 Golden 和 `format-positive` 负例均已按四页原尺寸渲染并逐页查看：Golden 仅将第 2 页 `-$15,250` / `-$15,500`、第 3 页 `−5.6%` / `(1.2%)`、第 4 页句中的 `-12` / `-3` 设为红色粗体，封面、蓝灰表格、正数和上下文均保持；无位移、裁切或文字断裂。
- 本题是精确格式修改与版式保持任务，因此使用 `reference`：视觉评分对照已核验 Golden 检查六处强调、正数误伤、符号/括号完整性、三张表结构与四页整体保持。字体 run、精确 RGB、目标字符、表格 OOXML、可编辑性、内容哈希和输出卫生仍由确定性 Judge 负责；VLM 补充肉眼可见的强调强度、字符衔接、溢出、对齐与整体专业度。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；对应综合分均为 `1.0000`。
- 中等可见负例选用本题已有的 `format-positive` 样例：六处负数、三张表和封面均保持，但第 3 页本应保持普通黑色的正数 `3.2%` 也被设为红色粗体；文件可正常打开，错误肉眼清楚且直接对应 V002，不属于空白或损坏交付。其确定性分 `0.0000`、视觉分 `0.8710`、综合分 `0.2613`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 6 次 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
