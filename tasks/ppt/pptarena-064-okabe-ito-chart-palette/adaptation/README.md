# PPTArena Case 64 逐题适配记录

## 适配契约

- 输入是 1 页带唯一原生簇状柱形图的演示文稿。
- 真实意图是只将 `North / East / West` 三个系列依次改为 `#0072B2 / #E69F00 / #009E73` Okabe–Ito 色盲友好配色。
- 主交付物为 `ColorblindPalette_GroundTruthA.pptx`。
- 系列名、顺序、Q1–Q4 分类、12 个数值、图表类型/标题/图例/坐标轴/标签/间距/几何和页面其他对象必须保持。
- 候选设计目标为 L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout。

## 上游来源与文件审计

- 来源：PPTArena Case 64: Colorblind-Safe Chart Palette。
- GitHub commit：`6f26a63ec1014b98c44990c7b58133b832070be2`；Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`；许可为 MIT。
- 本题直接使用上游原生可编辑 PPTX，不是本项目编造的 Mock 输入。

| 文件                                                      | 字节数 | SHA256                                                             | 实际检查                                                      |
|-----------------------------------------------------------|-------:|--------------------------------------------------------------------|---------------------------------------------------------------|
| `data/input_files/ColorblindPalette_TestA.pptx`           |  46185 | `a14d9499a9bea4738c648bfb686a25e2759edc753154cd439abbef9c8444d60e` | 1 页渲染；3 个对象、图表类型、系列/分类/数值、图例、坐标轴和 ZIP 成员审计 |
| 上游 `GroundTruth/ColorblindPalette_GroundTruthA.pptx` |  46417 | `527b1030dd1e23b940b20a2115b8bfb88e4f4753674836c243af3fd1d75a055e` | 三个配色真值正确；同时存在文档属性、缩略图、modId、图表唯一 ID 和视图属性改写 |
| `solution/Golden/ColorblindPalette_GroundTruthA.pptx`     |  43768 | `4b4f1db8a7b0407d9565824950d503e4baf21adac7eec9dd6d4704681ee66800` | 只改 `ppt/charts/chart1.xml` 中三个系列色，其他 39 个包成员字节不变         |
| `adaptation/upstream-case.json`                           |   1357 | `cf5fb14243e19c3b1f0e604cb62b000b99c70d778209fd65991505ee91514cbf` | Query、style target、enhancement note、路径和适配边界逐字段比对    |
| `tests/fixtures/golden-slide1.png`                        |  67362 | `edf4913ef1b7f5ade831874c52333a50bf2712eb9a05980a9e0aa07981790315` | 只用于整页压平负例，Judge 不将它作为答案或比对基准               |

## Golden 与 Judge 设计

- `build_golden.py` 锁定输入 SHA256，检查原色 `C83232 / 289646 / F08C1E` 均唯一出现，再在 ZIP 级精确替换为三个 Okabe–Ito 真值。
- 构建后对 40 个 ZIP 成员逐个计算哈希，确认只有 `chart1.xml` 改变；再用 python-pptx 回读三个系列色。
- Judge 不读 Golden，独立检查系列顺序与精确 RGB，并拒绝每个数据点的额外颜色覆盖。
- 图表语义签名覆盖类型、系列、分类、数值、标题、图例、数据标签、坐标轴和间距；页面标题和说明文本逐 run 保持。
- 独立 LibreOffice profile 渲染输入和候选，屏蔽图表区域后比较其他像素，并检查图表区可见。

## 实际验证结果

| 样例                                               | 实际 reward | 结果 | 关键原因                                   |
|----------------------------------------------------|------------:|------|--------------------------------------------|
| 本地 Golden                                       |    1.000000 | 通过 | 8 条判据全部通过                         |
| ZIP 压缩方式不同、Office 语义相同的合法变体      |    1.000000 | 通过 | 证明 Judge 不依赖 Golden 文件哈希              |
| fresh replay                                       |    1.000000 | 通过 | 空目录回放产物与已审计 Golden 哈希一致          |
| 错色序 / 非标准色 / 未修改                        |    0.000000 | 拒绝 | 精确配色门槛失败                        |
| 改数值 / 改系列名 / 改成水平条形图 / 删图例       |    0.000000 | 拒绝 | 图表数据或语义保持门槛失败                 |
| 重复图表 / 形状覆盖 / 移图表 / 改说明 / 整页压平 |    0.000000 | 拒绝 | 可编辑性、几何或非目标保持门槛失败             |
| 额外输出 `debug.txt`                          |    0.920000 | 部分 | 核心交付正确，但输出卫生正分失败                 |

同一 fresh replay 产物判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena064-oracle-okabe-ito-20260829`，trial 为 `pptarena-064-okabe-ito-chart-pal__yhVuJP6`，0 个异常；容器产物 SHA256 与本地 Golden 一致。

## 视觉检查与剩余风险

- 输入、上游 Ground Truth、本地 Golden 和合法变体均已渲染查看；三组新颜色、图例和四季度柱子均可读。
- `visual_qa.py` 对输入、Golden 和合法变体均报告 0 条警告。
- 正式 Release 前仍建议用 Microsoft PowerPoint 打开/保存抽检；L2 只是制题目标，正式难度必须经真实 Agent Rollout。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Case 与元数据、本题 `build_golden.py` / `build_fixtures.py` / `solve.py`、冻结输入和本地 Golden；核对了 North/East/West 系列、Q1–Q4、12 个值、原生簇状柱形图属性和只允许改三个系列填充色的边界。
- 冻结输入和本地 Golden 均已原尺寸渲染并逐页查看：输入使用红/绿/橙系列色；Golden 将 North、East、West 依次改为深蓝、橙黄、蓝绿色，并保持 Regional Sales 标题、三项图例、Q1–Q4、0–30 纵轴、网格、柱高、左侧 Instruction 和整页两栏布局。
- 本题有精确系列顺序、配色真值和冻结图表版式，因此使用 `reference`。专属视觉 Rubric 检查图例—系列颜色映射、色盲友好区分度、四组簇状柱结构、图表标题/图例/坐标轴排布以及非目标两栏页面保持。
- 精确十六进制色值、系列/季度/数值语义、原生可编辑图表、chart XML、图表几何和页面对象保持继续由确定性 Judge 负责；VLM 补充颜色辨识、过饱和/低对比感知、图例追踪、图表扫读和整页完成度。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；5 次确定性分与综合分也均为 `1.0000`。
- 中等可见负例使用本题已有的 `nonstandard-color` 样例：North 深蓝、East 橙黄、图表结构、数据、图例和两栏布局均保持，但将 West 从克制蓝绿色替换为高饱和荧光绿，仍可正常打开且全部图表文字可读。已原尺寸渲染并亲眼确认；其确定性分 `0.0000`、视觉分 `0.7650`、综合分 `0.2295`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 6 次 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
