# PPTArena Case 1 逐题适配记录

## 适配契约

- 输入是 18 页法语教学课件，混合法语、哈萨克语和少量英语，包含 4 张双语表格及多张背景图。
- 真实意图是只将指定哈萨克语译为英语，法语文字逐字保留，同时守住字体层级、表格可编辑性和 18 页结构。
- 主交付物是 `TranslatetoEnglishButKeepallFrenchGroundTruthA.pptx`。
- 候选设计目标为 L3 / Hard；正式难度字段保持 `unassigned`，等待真实 Agent Rollout。

## 上游来源与文件审计

- 来源：PPTArena Case 1: Translate to English But Keep French。
- GitHub commit：`6f26a63ec1014b98c44990c7b58133b832070be2`；Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`；许可为 MIT。
- 本题直接使用上游原始 18 页 PPTX、Query 和人工 Ground Truth，不是本项目编造的 Mock 数据。

| 文件                                                                                         |  字节数 | SHA256                                                             | 实际检查                                                                 |
|------------------------------------------------------------------------------------------------|--------:|--------------------------------------------------------------------|--------------------------------------------------------------------------|
| `data/input_files/TranslateToEnglishButKeepallFrenchTestA_Bonjour mes amis!.pptx`              | 3873562 | `166e95dcf7883aa0ed60768e184c225851161a908f82223a7f8443f45cdd32dd` | 18 页逐页渲染；逐对象读取文字、run 语言/样式、表格、几何、图片与备注     |
| `adaptation/upstream-ground-truth.pptx`                                                        | 3871749 | `32a2adb5d145439a959d53b7dc2f24dabb54d01e37cc0cbd2c0a9b3c34dde7c2` | 逐页查看上游人工译文；发现四处明确语义问题和两个法语框的 run 重写          |
| `solution/Golden/TranslatetoEnglishButKeepallFrenchGroundTruthA.pptx`                           | 3729675 | `63cc22f03dbcd0e645e95b7de03e0a11812adc79120df9a4343604de5c73e2e7` | 校正四处译文，恢复第 10、12 页法语对象原 OOXML；18 页视觉与结构复核          |
| `adaptation/upstream-case.json`                                                               |    1612 | `93c09b84a1ff073617ba598068a2d07a7c7734d0d9ec7ac220fe5a01c5c0e9f7` | Query、style target、路径、类别与 enhancement note 逐字段比对          |
| `tests/fixtures/golden-slide9.png`                                                            |   81568 | `d90e6d0f257306106180ac9d5ac693667bdb94282a602cc61257ddea32b8e01d` | 只用于第 9 页表格压平负例，Judge 不将其当作答案                       |

## 上游答案审计与 Golden 决策

上游人工 Ground Truth 完成了主要翻译，但逐页核对后确认以下四处不能直接进入 Benchmark 真值：

- 第 4 页 `tamakh, azyq-tulik` 的 `meal, a diet` 改为更贴合原意的 `food, groceries`。
- 第 6 页哈萨克语语法词 `zhiktelui` 在“动词及其……”上下文中应为 `conjugation`，不是 `classification`。
- 第 8 页右列 `kofe` 应为饮品 `Coffee`，不是场所 `Cafe`。
- 第 12 页 `Su ishu` 包含“水”，因此从 `To drink` 校正为 `To drink water`。

上游还重新合并了第 10、12 页左侧法语文本的 run。虽然文字、字号和字体值看似相同，第 10 页已出现字距渲染差异。本地 `build_golden.py` 因此恢复这两个法语对象的原始 OOXML，再保留右侧英译。

## Judge 设计

- Judge 不读取 Golden。它从冻结输入独立定位 8 个目标文本对象与 4 张表格的 40 个右列单元格，逐项检查已审定英译。
- 非目标文本对象、法语表格左列和所有备注逐对象比对文字与可见样式；媒体集合按 SHA256 比对。
- 目标译文的 Times New Roman、字号、粗斜体、红色标题和表格样式分层检查；允许目标文本框为可读性小幅调整。
- 表格行列、填充、边框和结构在排除 volatile `modId/rowId` 后比对。
- 独立 LibreOffice profile 渲染 18 页；遮罩译文可能重排的目标区域后，其余页面像素必须与输入保持。

## 实际验证结果

| 样例                                             | 实际 reward | 结果 | 关键原因                                         |
|--------------------------------------------------|------------:|------|--------------------------------------------------|
| 本地复核 Golden                                  |    1.000000 | 通过 | 9 条判据全部通过                                |
| Office 语义不变、ZIP 压缩方式改变的合法变体 |    1.000000 | 通过 | 证明 Judge 不依赖 Golden 文件哈希                 |
| fresh replay                                     |    1.000000 | 通过 | 空目录回放产物哈希与 Golden 一致                  |
| 未修改 / 漏译 Coffee / 上游未纠错答案        |    0.000000 | 拒绝 | 目标文本或 40 个表格译文门槛失败                  |
| 把 conjugation 写成 classification / 漏 water |    0.000000 | 拒绝 | 语义正确性门槛失败                                |
| 把法语 `Le pain` 翻成英语                     |    0.000000 | 拒绝 | 法语左列逐单元格保护失败                          |
| 改标题颜色 / 移文本框 / 删对象 / 压平表格  |    0.000000 | 拒绝 | 样式、几何、结构或原生可编辑性失败                 |
| 夹带答案路径                                     |    0.000000 | 拒绝 | 对象结构和答案标记失败                            |
| 额外输出 `debug.txt`                              |    0.960000 | 部分 | 主交付正确，但输出卫生正分失败                     |

同一 fresh replay 产物判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena001-oracle-kazakh-french-20260829`，trial 为 `pptarena-001-kazakh-to-english-p__Xeqz5uW`，0 个异常；容器产物 SHA256 与本地 Golden 一致。

## 视觉检查与剩余风险

- 输入、上游 Ground Truth、本地 Golden 和合法变体的 18/18 页均已渲染查看；校正后译文可读，表格无裁切，法语保留。
- `visual_qa.py` 对输入、Golden 和合法变体均报告第 1、4、6 页原有文本框稀疏及第 12 页原有稀疏提示；Golden 第 6 页文本框扩高是为了译文完整可读，不是内容丢失。
- 正式 Release 前仍需对课件中作者姓名字样、背景图和第三方教学内容做项目级法务与隐私复核；L3 只是制题目标，正式难度必须经真实 Agent Rollout。
