# PPTArena Case 61 逐题适配记录

## 适配契约

- 输入是 1 页带 6×3 原生表格和 `Sample Signal` 图片的演示文稿。
- 真实意图是保留表头，将 5 个完整数据行按 `Score` 从高到低排序。
- 主交付物为 `SortByScoreAndCropImage169_GroundTruth.pptx`。
- 不可妥协约束：`Bob/B/96 -> Eve/B/91 -> Diego/A/88 -> Alice/A/82 -> Cara/C/74`；行关联完整；表格样式不变；标题不变；非目标图片的像素、几何和裁剪不变。
- 候选设计目标为 L2 / Medium，正式难度字段保持 `unassigned`，等待真实 Agent Rollout。

## 上游来源与许可

- 来源：PPTArena Case 61: Sort By Score And Crop Image 169。
- GitHub commit：`6f26a63ec1014b98c44990c7b58133b832070be2`；Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 和 Hugging Face 数据卡均标记 MIT，副本已保存在本目录。
- 上游 Query、style target、enhancement note 和路径已逐字段保存在 `upstream-case.json`。
- 本题直接使用上游真实 PPTX，不是本项目编造的 Mock 输入。

## 已读入上下文的文件

| 文件                                                                  | 字节数 | SHA256                                                             | 实际检查                                                     |
|-----------------------------------------------------------------------|-------:|--------------------------------------------------------------------|--------------------------------------------------------------|
| `data/input_files/SortByScoreAndCropImage169_Test.pptx`               |  94508 | `dd2ee0141319451ea8b33b20c06b49e2985932f302ba6de1349566a864f7995c` | 1 页渲染；3 个对象、6×3 表格、全部单元格、图片裁剪和媒体审计       |
| 上游 `GroundTruth/SortByScoreAndCropImage169_GroundTruth.pptx`    |  94351 | `80d8cfac1e5395e7d8e11f7075b022d4773c541f655879ac64cccaa2e10828e1` | 1 页渲染；排序真值和非目标图片上下各 12.5% 裁剪/移动审计          |
| `solution/Golden/SortByScoreAndCropImage169_GroundTruth.pptx`         |  87647 | `1c89355c12156d06eb4cd71d81dfc444767ed8dcee764b2a7078aa22ac5c16cf` | 1 页渲染；只有表格数据行重排，行样式、标题和图片全部保持审计          |
| `adaptation/upstream-case.json`                                      |   1574 | `47d9732446216fa4ad346de363e9e8db3fb0578395d40b28f33e02882ee87ed1` | Query、style target、enhancement note、路径和适配边界逐字段比对   |
| `tests/fixtures/golden-slide1.png`                                   | 110917 | `8b16639774cbb15a1566d8a12d73f4daa5fec4c97dabfa1be6d19ef6a96ef38b` | 只用于构造整页压平负例，Judge 不把它当作答案或比对基准            |

## Golden 构建与上游审计

- `build_golden.py` 先锁定输入 SHA256，再从源表格中深拷贝每个单元格的原生 `txBody`，按分数降序将完整行移到目标位置。
- 表格几何、行高、列宽、`tblPr`、18 个单元格 `tcPr`、标题和图片签名在保存前后均独立比对。
- 上游 Ground Truth 的表格排序正确，但它额外将 `Sample Signal` 图片上下各裁剪 12.5% 并轻微左移/上移，导致图片自带标题和 x 轴标签被截掉。
- 该图片修改没有出现在 Query 中，而且明显降低可读性；因此本地 Golden 明确不复制这一上游脏改动。

## Judge 设计

本地 Judge 不读 Golden，只根据冻结输入和任务真值独立重算。

- 要求唯一原生 6×3 表格，表头不动，数据行必须严格为 96、91、88、82、74 的降序顺序。
- 用多重集检查五个 Name/Team/Score 三元组，拒绝错配、重复、丢失和额外行。
- 比较表格结构与单元格底层样式，同时按逐字符样式而非 run 数量验证文字格式，因此合法 run 分段变体可以通过。
- 图片必须保持底层 blob SHA256、位置、尺寸与四边裁剪；上游式裁剪和移动会直接失败。
- 用独立 LibreOffice profile 渲染候选和输入，屏蔽表格区域后比较其他像素，并检查表格文字可见。

## 实际验证结果

| 样例                                             | 实际 reward | 结果 | 关键原因                                |
|--------------------------------------------------|------------:|------|-----------------------------------------|
| 本地 Golden                                     |    1.000000 | 通过 | 9 条判据全部通过                      |
| Bob 文本 run 分段不同的合法变体                |    1.000000 | 通过 | 文本、逐字符样式、数据、保持与渲染合格       |
| fresh replay                                     |    1.000000 | 通过 | 空目录回放产物与已审计 Golden 哈希一致         |
| 空交付 / 损坏 PPTX / 未排序                  |    0.000000 | 拒绝 | 缺文件、无法打开或排序/样式门槛失败            |
| 错顺序 / 行错配 / 表头移动 / 重复行           |    0.000000 | 拒绝 | 排序或行关联门槛失败                   |
| 破坏单元格填充 / 改标题 / 上游式裁图 / 删图 |    0.000000 | 拒绝 | 样式或保持门槛失败                     |
| 文本框伪表格 / 整页截图 / 答案标记            |    0.000000 | 拒绝 | 可编辑性或答案泄漏门槛失败                |
| 额外输出 `debug.txt`                        |    0.940000 | 部分 | 核心交付正确，但输出卫生正分失败                |

同一 fresh replay 产物独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena061-oracle-sort-score-20260829`，trial 为 `pptarena-061-sort-table-by-score__pbJ4UUt`，0 个异常；容器产物 SHA256 为 `1c89355c12156d06eb4cd71d81dfc444767ed8dcee764b2a7078aa22ac5c16cf`，与已审计 Golden 一致。

## 视觉检查与剩余风险

- 输入、上游 Ground Truth、本地 Golden 和合法变体均渲染并人工查看；本地 Golden 表格排序正确，图片标题和坐标轴完整。
- `visual_qa.py` 对输入、本地 Golden 和合法变体均报告 0 条启发式警告。
- 正式 Release 前仍建议用 Microsoft PowerPoint 打开/保存抽检；L2 只是制题配比目标，正式难度必须经真实 Agent Rollout。
