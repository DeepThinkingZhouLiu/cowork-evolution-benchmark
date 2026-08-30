# PPTX 单题质量门槛

## 目录合同

每题至少包含：

- `instruction.md`
- `task.toml`
- `rubric.json`
- `environment/Dockerfile`
- `environment/input_files/` 或等价的镜像内只读输入
- `data/input_files/`
- `solution/solve.sh`、任务专属源文件和可打开的 Golden PPTX
- `tests/test.sh`、`tests/judge.py`、`tests/expectations.json`
- `metadata/task.json`、`metadata/provenance.json`、`metadata/input-manifest.json`
- `adaptation/README.md`、上游任务/验收依据副本和必要的许可/NOTICE

`task.toml` 必须使用项目当前 Schema；网络默认关闭；Agent、Verifier 和 artifacts 路径不得重叠或泄漏 Golden。Verifier 必须写 `/logs/verifier/reward.txt`，详细报告可另外写入 `/logs/artifacts/`。

## Golden 门槛

- 所有事实可回溯到冻结输入或明确允许的计算；不得编造业务数据。
- 叙事顺序、受众、结论和行动建议与原任务一致。
- 文本、形状、表格和图表保持可编辑；必要图片保留来源和用途。
- 字体、字号、对比度、安全边距、页码、页脚、来源和模板约束可读。
- 没有占位符、溢出、遮挡、离页对象、损坏关系、错误链接或空白页。
- 用 LibreOffice 打开并渲染全部页面；查看联系表后，再逐页查看原尺寸图。
- 对复杂图表、流程、动画、母版、备注或嵌入对象，同时检查 OOXML 结构。

## Rubric 门槛

- 至少 6 条原子判据；事实正确性、核心业务目标和文件完整性权重最高。
- 至少一个不可被其他优点抵消的 hurdle；关键违规使用 penalty。
- 分开评估：文件完整、内容/事实、结构/叙事、视觉可读、原生可编辑、约束遵循；有跨应用状态时再加状态一致性。
- 每条判据写明可观察证据和允许容差。避免“美观”“专业”等不可复现的空泛判据。
- 不把页数、文件名、主题色或关键词存在性当作主要质量分。
- 不要求候选文件与 Golden 的对象 ID、压缩字节、内部顺序或哈希一致。

## Judge 门槛

- 从候选 PPTX 和冻结输入独立重算真值；对数字、日期、实体、图表数据和跨文件关系做语义检查。
- 检查 PPTX 可打开、页数、关键对象可编辑、文字未被整页截图替代、图片/图表关系有效。
- 对布局至少检查离页、重叠、过小字体、文字溢出、对比度和关键对象遮挡；复杂视觉项可组合确定性检查与盲评，但随机评审不能主导核心事实分。
- 对跨应用题从只读 sidecar 日志或收集钩子取得状态证据，Agent 不得直接覆盖这些证据。
- 每个判据在结果中返回分数、是否通过和证据；总分公式、hurdle 截断和 penalty 必须与 `rubric.json` 一致。
- 异常、超时、缺文件、损坏压缩包和依赖失败默认保守给分，不得误接受。

## 必测样例

- Golden 正例：应满分并通过全部 hurdle。
- Fresh replay：从空输出目录运行 `solve.sh` 后再评分，应与 Golden 一致。
- 空交付：必须为 0。
- 损坏/伪装文件：必须为 0。
- 未修改输入原件：必须失败。
- 事实错误：修改一个关键数值、日期或实体，相关高权重项必须失败。
- 不完整：删除一页、关键结论、来源或状态操作，分数显著下降。
- 压平作弊：把页面替换为截图，editability 项和相关 penalty 必须触发。
- 布局错误：制造遮挡、离页或不可读字体，视觉项必须下降。
- 中等样例：只完成主要内容但存在真实缺陷，分数应严格位于空交付与 Golden 之间。
- 硬编码/泄漏：输出夹带答案标记、solution 路径或 Judge 关键词不能提高分数。

## 单题证据记录

`adaptation/README.md` 至少记录：来源路径和 commit、查看过的每个输入、输入/Golden SHA256、上游缺陷、适配边界、Golden 制作判断、每个负例及预期、实际回放命令与分数、五次重复评分、渲染页数、逐页视觉结论和剩余风险。
