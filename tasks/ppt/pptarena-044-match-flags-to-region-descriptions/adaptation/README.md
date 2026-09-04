# PPTArena Case 44 逐题适配记录

## 适配契约

- 角色：演示文稿内容与图片排版编辑。
- 输入：1 页 4:3 非洲地区联络幻灯片，页面底部实际叠放 5 面国旗。
- 真实意图：删除不在地区清单中的隐藏南非国旗，将肯尼亚、尼日利亚、埃及和赞比亚国旗分别与东非、西非、北非和南部非洲机构说明配对。
- 主交付物：`MatchFlagtoText_GroundTruthA.pptx`。
- 不可妥协约束：4 面正确国旗各一次；南非国旗删除；4 条说明完整且独立；配对正确、距离近、无重叠与越界；主标题和顶部两图保持。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 44: Match Flag to Text。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>，commit `6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；副本保存在本目录。
- 上游 Query、style target、enhancement note、输入和人工 Ground Truth 路径已原样保存在 `upstream-case.json`。
- 已完整阅读上游 `src/evaluation_pairs_refined.json`、`src/llm/judge.py` 和 `agent_bench/judge_predictions.py`；未运行依赖外部模型的网络 Judge。
- 本题沿用上游真实 PPTX、5 面原始国旗和 4 条原始地区机构说明，不是本项目自行编造的 Mock 数据。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/MatchFlagtoText_TestA.pptx` | 180299 | `8c2c6f409684507f2d56cc205c8a6b17d69e65d5036748dbe6d2098e24f325e2` | 1 页渲染；11 个页面对象、7 个媒体、国旗叠放、文字、几何和备注审计 |
| 上游 `GroundTruth/MatchFlagtoText_GroundTruthA.pptx` | 168222 | `2058a3cbab5c9267caaafb2cae133ebaf234b9356552d0d11fa22104f66d4ff6` | 1 页渲染；国旗/文字映射、越界、重叠、标题顶图可见性和包差异审计 |
| `solution/Golden/MatchFlagtoText_GroundTruthA.pptx` | 180299 | `2b34894adfd13a72970accb30d212b3239e6672095a7db079239dd9f53b09067` | 本地 2×2 排版；原生文字、媒体哈希、几何、PDF 文本和顶部像素审计 |
| `adaptation/upstream-case.json` | 1935 | `d07221a34b55e95c99b9efc17b04ebeb8719694ac846f90e8d85a1864bc5f1e4` | Query、style target、enhancement note、路径和改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、评测方式和 MIT 标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开状态和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本 |
| `tests/fixtures/golden-slide1.png` | 300846 | `78b892570cc3a2955c38f6170716108ce0a698b23f1722b4a74bbf73f002ff89` | 仅用于构造整页截图压平负例 |

## 国旗真值与媒体审计

| 对象 | 国家 | 对应说明 | 图片 SHA256 | 处理 |
|------|------|----------|---------------|------|
| 原 shape id 1030 | 肯尼亚 | `EAST AFRICA - Kenyan Competition Authority` | `f6ad8649736281f1d3d1d6093e8f2654251a0a8b061cd4ced45c895374bc0188` | 保留并与东非说明配对 |
| 原 shape id 7 | 尼日利亚 | `WEST AFRICA - Nigerian Consumer Protection Council` | `021690c3c58a55242b53d43e07bb0c0daf4d0f1bd70aabb12a2feb1c7757bb1a` | 保留并与西非说明配对 |
| 原 shape id 1032 | 埃及 | `NORTH AFRICA - Egyptian Consumer Protection Agency` | `8b98a3cf1c84aa7aac4a45637b7da811f38ea7d1ba2ea1d0c8f27310b107d80e` | 保留并与北非说明配对 |
| 原 shape id 1026 | 赞比亚 | `SOUTHERN AFRICA - Zambian Competition and Consumer Protection Commission` | `128d4315add1153b7be16b728dc578f5d5f955f9ac245e428fda0f2adac4c8a8` | 保留并与南部非洲说明配对 |
| 原 shape id 1034 | 南非 | 无对应地区条目 | `4f1abe2384af08da1d2113c8395095411f0e8b86cfb7db9f154d1137844af5b0` | 删除 |

南非国旗与赞比亚国旗的原始几何完全一致，均为 x=7.5、y=5.5085、w=2.5、h=1.9915 英寸；赞比亚国旗在 z-order 上层，因此原渲染只能看到四面国旗。

## Golden 构建与上游答案审计

- `build_golden.py` 只处理本题锁定哈希的输入；它使用上表的 5 个图片哈希定位对象，删除原列表和南非国旗，添加 4 个原生文字对象。
- 本地 Golden 采用两列两行：每面国旗为 1.75×1.40 英寸，文字位于对应国旗右侧 0.25 英寸；8 个配对对象全部位于页内且无重叠。
- 主标题、左上非洲地图和右上 FTC 徽章的对象、几何、文本样式和媒体哈希与冻结输入一致。
- 独立合法变体采用四行水平排版，不复制 Golden 的几何，但国旗、文字、配对、顶部保持和可编辑性全部合格。
- 上游 Ground Truth 并不符合本地更严格的可发布契约：北非文本框横跨西非/北非区域，按边界距离更接近尼日利亚国旗；部分文本框 x 为负数；赞比亚文本与旗帜发生 0.4915 平方英寸重叠，西非与北非对象也存在 0.8361 平方英寸跨组重叠。
- 上游 Ground Truth 在独立 LibreOffice profile 中的主标题可从 PDF 提取，但像素渲染不显示标题和两张顶图；本地 Judge 因映射歧义、越界与重叠门槛将其拒绝为 0。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只使用冻结输入、任务契约和上表的真实媒体映射。

- 按图片 blob SHA256 定位 4 面正确国旗、南非国旗和 2 张顶图；正确国旗必须各一次，南非国旗必须无页面引用。
- 要求 4 条说明分别精确存在于 4 个独立文字对象；一个文本框同时包含多条说明会被视为原列表未拆分。
- 对每个说明文本框计算到 4 面国旗边界矩形的距离；最近国旗必须是正确国家，正确配对距离不得超过 0.8 英寸。
- 逐对检查 8 个配对对象的页内边界和两两重叠；国旗必须是原图片对象，文字必须是可编辑文本对象，拒绝整页截图。
- 主标题与冻结输入比较页面对象类型、占位符、几何、文本、段落、run 字体样式和文本框属性；顶图比较媒体哈希与几何。
- 使用独立 LibreOffice profile 转 PDF，校验标题和四组机构关键词可提取、4 个文本框区域实际有深色像素，并检查顶部 19.5% 区域与冻结输入的像素相似度不低于 0.995。
- 排版质量正分检查四面国旗面积比不高于 1.25、单面面积在 0.7–4.5 平方英寸且配对间距标准差不高于 0.30 英寸。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地 2×2 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 独立四行配对合法答案 | 1.000000 | 通过 | 国旗、文字、近邻配对、顶部保持和可编辑性正确 |
| fresh replay | 1.000000 | 通过 | 空输出目录回放产物 SHA 与已审计 Golden 一致 |
| 上游人工 Ground Truth | 0.000000 | 拒绝 | 北非说明距离尼日利亚国旗更近，并有越界和跨组重叠 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺失、无法打开或 5 面国旗/整段列表仍存在 |
| 南非国旗残留 / 缺国旗 / 重复国旗 | 0.000000 | 拒绝 | 正确国旗唯一性或额外国旗门槛失败 |
| 调换说明 / 截断说明 / 把四条重新合并为列表 | 0.000000 | 拒绝 | 近邻国旗映射、原文或独立对象门槛失败 |
| 改标题 / 删顶图 / 移顶图 | 0.000000 | 拒绝 | 不可修改区域保持失败 |
| 文本越界 / 两组重叠 / 整页截图 / 额外加页 | 0.000000 | 拒绝 | 页内边界、重叠、可编辑性或页数门槛失败 |
| 夹带 `solution/Golden` | 0.780000 | 不通过 | 输出卫生失分且触发严重错误扣分 |
| 单面国旗明显缩小但仍正确配对 | 0.930000 | 中等 | 核心真值正确，仅国旗尺寸协调性失分 |

同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena044-oracle-match-flags-20260829`，trial 为 `pptarena-044-match-flags-to-regi__DevyCaj`，0 个异常；容器产物 SHA256 为 `2b34894adfd13a72970accb30d212b3239e6672095a7db079239dd9f53b09067`，与已审计 Golden 一致。

## 视觉检查与剩余风险

- 冻结输入、本地 Golden、四行合法变体和上游 Ground Truth 均已用独立 LibreOffice profile 原尺寸渲染并人工检查。
- `visual_qa.py` 对输入和本地 Golden 都只报告同一条上游既有的标题 `underfilled_textbox` 通用警告；本题不擅自改动标题占位符。
- 本机默认共享 LibreOffice profile 在密集连续转换时偶发返回旧布局缓存；Judge 每次创建独立 profile，合法四行变体也已在独立 profile 中人工确认标题、顶图和四组配对都存在。
- 正式 Release 前仍建议用 Microsoft PowerPoint 打开/保存抽检；L2 / Medium 只是制题配比目标，正式难度必须经过冻结 Baseline Agent 重复 Rollout 后再写回。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本改造记录、上游 Case、本题 `build_golden.py` / `build_fixtures.py` / `solve.py`、冻结输入和本地 Golden，并核对了 5 面上游国旗的国家真值、4 条地区机构原文、图片哈希、最近邻配对门槛和页眉冻结边界。
- 冻结输入与本地 Golden 均已原尺寸渲染并亲眼对照：输入页把四条说明留在项目符号列表中，底部国旗与说明没有形成组；Golden 删除隐藏南非国旗，将肯尼亚—东非、尼日利亚—西非、埃及—北非、赞比亚—南部非洲排成清楚的 2×2 单元，并完整保留主标题、非洲地图和 FTC 徽章。
- 本题属于有明确国家真值与已核验目标页的版式修复，因此使用 `reference`；Rubric 明确允许四行或其他无歧义的合法布局，不要求候选复制 Golden 的 2×2 几何。专属视觉标准分别检查四组配对的感知清晰度、组内/组间排版、四条长短不同说明的真实可读性、四面国旗的比例与视觉重量、以及冻结页眉到主体的专业连贯性。
- 国旗图片 blob、国家唯一性、南非国旗删除、说明原文、最近邻映射、对象可编辑性、页内边界、顶部对象签名和包完整性继续由确定性 Judge 负责；VLM 只补充肉眼可见的配对歧义、留白/对齐、换行可读性、国旗失真和整页专业完成度。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；5 次综合分均为 `1.0000`。
- 中等可见负例直接使用本题已有的 `inconsistent-flag-size` 边界样例：四个国家、4 条原文、2×2 配对、页眉和对象可编辑性均保留，但肯尼亚国旗从 1.75×1.40 英寸缩至 1.25×1.00 英寸，同时把其说明左移，使左上单元的旗帜视觉重量和组内基线明显弱于其余三组。已原尺寸渲染并亲眼确认文稿可打开、全部内容可读且缺陷肉眼可见；该负例确定性分 `0.9300`、视觉分 `0.8050`、综合分 `0.8925`，严格低于 Golden 均值且高于空白/损坏交付。
- Golden 第 3 次首次调用和负例首次调用曾因 Codex/Provider 非零退出被标记为评测基础设施失败；二者均已在独立结果目录成功重试，失败记录未计 0 分、也未混入稳定性统计，当前没有未解决基础设施失败。
