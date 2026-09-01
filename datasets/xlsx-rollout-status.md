# Excel Agent Rollout 状态

## 固定配置

| 配置 | Harness | 模型 | Provider | 状态 |
|---|---|---|---|---|
| `codex-gpt56sol` | Codex CLI `0.150.1` | `gpt-5.6-sol` | ZhiCloud Responses | 配置完成，Harbor Agent 安装阻塞 |
| `opencode-opus48` | OpenCode `1.18.25` | `claude-opus-4-8` | ZhiCloud Anthropic Messages | 阻塞 |

两组同时改变了 Harness 和模型，因此只作为交叉复核，不能把分数差异单独归因于 Harness 或模型。

## Codex Harbor 烟测（2026-08-28）

`officeval-090-multi-sheet-summary-extraction` 的四次冷启动均停在 Codex Agent 安装阶段，尚未执行用户任务或 Judge：前三次依次暴露 Debian 软件源和 `raw.githubusercontent.com` 访问限制；加入精确环境白名单后，第四次成功下载 NVM 安装脚本，但 `git clone nvm` 遇到 HTTP/2 framing error。旧烟测曾传入后来废弃的 solver `xlsx` Skill；这些运行只保留为基础设施故障记录，不属于当前无辅助评测协议，也不计为模型得分。

正式难度 Rollout 不向 Codex 或 OpenCode 挂载任何 Skill、Memory、few-shot 或制题材料。两组仅使用各自原生 Harness、任务可见指令、输入和环境工具。

## 第二批无辅助基础设施预检（2026-08-29）

对 C007 新建了符合冻结协议的 Harbor 0.21.0 安装预检 `xlsx-batch2-codex-install-smoke-20260829`：`codex-cli@0.150.1`、`gpt-5.6-sol`、`skills=[]`、无 Memory/few-shot/制题材料，且只执行 `install_only`，不运行用户任务或 Judge。任务镜像构建完成后，Agent 安装始终停在 `apt-get update && apt-get install -y curl bash nodejs npm ripgrep`，360 秒后由 Harbor 报 `AgentSetupTimeoutError`。记录中 `agent_execution=null`、`verifier=null`、Token/成本为空；Harbor 汇总显示的 `Mean 0.000` 只是异常 trial 的占位聚合，禁止解释为模型分数。

本次证据保存在 `.coworkevo/runs/xlsx-batch2-codex-install-smoke-20260829/`。与 2026-08-28 的旧烟测不同，本次明确满足 `skills=[]`；但仍未进入模型执行，因此不展开 4 题 × 3 次正式 Codex rollout。

## OpenCode 阻塞证据（2026-08-28）

- ZhiCloud `GET /v1/models` 只返回 `gpt-5.5`、`gpt-5.6-sol`、`gpt-5.6-terra`。
- 使用所提供 key 对 `claude-opus-4-8`、`claude-opus-4.8`、`opus-4.8` 和 `claude-opus-4-6` 发起最小 Anthropic Messages 请求，均返回 HTTP 503：当前“GPT 特惠渠道”分组无可用 Claude distributor。
- OpenCode 配置已使用 `{env:ZHICLOUD_API_KEY}`，仓库中未保存密钥。获得带 Claude/Opus 渠道的 key 或调整账户分组后，才能执行真实 Rollout。

2026-08-29 再次复核：`GET /v1/models` 仍只返回 `gpt-5.5`、`gpt-5.6-sol`、`gpt-5.6-terra`；对冻结模型 `claude-opus-4-8` 的最小 Anthropic Messages 请求仍返回 HTTP 503，错误为当前“GPT 特惠渠道”分组无可用 distributor。未启动 OpenCode Agent，不产生模型分数。

## 运行记录要求

每次运行记录任务 commit、Harness/CLI 版本、模型 ID、`skills=[]`、Prompt digest、预算、随机种子、产物、逐条 Rubric、总分、成功与否、失败类型、步骤数、Token 和成本。未实际运行的格子保持空值，禁止填充推测分数。

## 当前制题准备进度（2026-08-29）

本阶段目标是从已审计的公开 Excel 候选中，逐题制作与现有 14 道 Excel 基线相同的 Harbor 任务包。候选题不是正式 Release，也不是把上游题目原样复制进数据集；每题都必须经过来源核验、输入恢复、真值独立确认、离线化、Golden/Judge 验证和人工复核。

| 工作项 | 状态 | 说明 |
|---|---|---|
| 现有 14 道 Excel 基线人工审计 | 已完成 | 作为独立 baseline 保存，不与本轮候选混合；人工审计结果见 `xlsx-baseline-audit.csv`。 |
| Excel 上游候选筛选 | 已完成 | 共 97 道：`ready=70`、`conditional=27`；逐题清单见 `xlsx-candidate-inventory.csv`。 |
| 来源、版本、许可和可用性审计 | 已完成 | 审计结论见 `reference-source-audit.md`；`conditional` 候选在许可证或离线化缺口关闭前不得发布。 |
| Excel Harbor 制题 Skill | 已完成 | `.agents/authoring-skills/author-xlsx-harbor-tasks/SKILL.md`；只用于作者侧制题，不注入被测 Agent。 |
| 新候选转换为 Harbor 任务包 | 九个小批次已完成 | pilot C001/C008/C013/C096、第二批 C007/C020/C085/C093、第三批 C002/C003/C005/C011、第四批 C006/C010/C012、第五批 C015/C017/C018、第六批 C082/C014/C022、第七批 C095、第八批 C031/C016/C045/C060/C061、第九批新增 19 道，共 46 道候选；本轮收口到 60 道，不再继续扩展剩余候选。 |
| Golden、Judge、负例和渲染验证 | 九个小批次已完成 | 新增 46 题均完成 Golden、fresh replay、正负例和原生格式检查；第九批 19 题的 Golden/fresh replay 均 `1.000000`，空/损坏/未修改输入均不通过。全库公开端到端静态校验通过 `103/103`，详见各题 `adaptation/README.md`。 |
| Codex/OpenCode 无辅助难度 rollout | 基础设施阻塞 | 2026-08-29 无 Skill Codex 预检仍停在 Agent 安装；OpenCode Opus 渠道仍返回 503。没有可填入的模型分数。 |
| 人工逐题复核 | 待 rollout 后 | 由项目负责人决定接受、修改、降级或淘汰；Codex 不得代替该决定。 |

### 候选难度假设

候选清单中的 L1–L4 是制题前的结构性假设，不是最终标签：

- **L1：单应用基础操作**，例如单工作簿中的直接修改、筛选、排序或简单汇总。
- **L2：单应用复杂操作**，例如多阶段公式/模型、图表、模板保持、宏或多个结果区域。
- **L3：跨应用操作**，例如 Excel 与 Word、邮件、日历或 PDF 之间的连贯转换和交付。
- **L4：跨应用长程任务**，例如多个异构来源、状态变化或冲突处理、多个交付物和多个依赖检查点。

初步分布为 L1=9、L2=36、L3=37、L4=15。该分布只用于选择 pilot 和安排制题顺序。最终等级必须依据不挂载任何 solver Skill、Memory、few-shot 或制题资料的 Codex/OpenCode rollout，并结合成功率、Judge 分数、失败类型、步骤数、成本和人工复核重新校准。

### Harbor 制题阶段的标准产物

每道通过作者侧验证的候选题应至少包含：

```text
tasks/xlsx/<task-id>/
├── instruction.md              # Agent 可见的用户任务和交付要求
├── task.toml                   # Harbor 运行配置
├── rubric.json                 # 隐藏侧原子评分标准
├── environment/Dockerfile      # 固定、离线运行环境
├── data/input_files/           # 作者侧输入副本
├── solution/solve.py           # 任务专用参考解法
├── solution/solve.sh           # Golden 生成入口
├── solution/<Golden files>     # 隐藏侧参考交付物
├── tests/judge.py              # 候选产物 Judge
├── tests/test.sh               # 标准评分入口
├── tests/expectations.json     # 可选的确定性真值
├── metadata/task.json          # 任务和初步分类元数据
├── metadata/input-manifest.json
├── metadata/provenance.json    # 来源、许可、版本和 SHA256
└── adaptation/README.md        # 证据、改造、验证和待人工决策
```

`instruction.md` 只能暴露求解所需的任务背景、输入用途、输出文件名、保存位置和专业要求；Golden、完整 Rubric、Judge 实现、作者记录和难度分区属于隐藏评测侧。制题阶段可以由 Codex 按 Skill 编写任务专用脚本和检查器，但评测阶段的 Codex/OpenCode 不得读取这些辅助文件。

### 下一阶段里程碑

1. 关闭 Codex Harbor Agent 安装超时和 OpenCode Opus provider 缺口；修复必须属于评测基础设施，不得通过向 solver 注入 Skill、作者脚本或隐藏文件绕过。
2. 基础设施预检通过后，对已完成的 8 题按冻结矩阵分别运行无辅助 Codex/OpenCode，每题 3 次，记录 `skills=[]`、版本、模型、Prompt digest、产物、逐条 Rubric、总分和失败分类。
3. 将 Agent 安装、环境构建、Provider/API 故障与实际模型失败分开；任何未进入 `agent_execution` 的 trial 都不得产生模型分数。
4. 由项目负责人结合作者侧验证和真实 rollout，对 8 题逐题决定接受、修改、降级或淘汰，并重新校准 L1--L4。
5. 根据已观察的通过率与失败类型，从剩余 `ready` 候选中定向选择下一批 3--5 题；`conditional` 候选继续暂缓。
6. 只有多个滚动批次证明制题和 rollout 流程稳定后，才决定是否扩展到正式 60 道；来源、输入哈希、模板和语义近重复按 leakage group 隔离后，才进入 `feedback / selection / final` 切分。

当前阶段的完成定义是：pilot 任务包达到 `pending_human_review`，而不是自动标记为已接受或正式发布。

### Pilot 制题结果（2026-08-29）

| 候选 | Harbor 包 | 作者假设 | Golden/Judge | 负例与渲染 | 当前状态 |
|---|---|---|---|---|---|
| XLSX-C001 | `tasks/xlsx/officebench-1-4-score-update` | L1 | Golden、fresh replay、Judge=1.000000 | 空/损坏/未修改；LibreOffice PDF | `pending_human_review` |
| XLSX-C008 | `tasks/xlsx/officebench-1-6-class-schedule` | L2 | Golden、fresh replay、Judge=1.000000 | 空/损坏/未修改；LibreOffice PDF | `pending_human_review` |
| XLSX-C013 | `tasks/xlsx/officebench-2-13-section-calendar` | L3 | Golden、fresh replay、Judge=1.000000 | 空/缺员/Bob/错误时间；结构检查 | `pending_human_review` |
| XLSX-C096 | `tasks/xlsx/cowork-bench-supplier-contract-compliance` | L4 | Golden、fresh replay、Judge=1.000000 | 空/缺交付/错分/跨文件/损坏；LibreOffice PDF | `pending_human_review` |

以上分数只表示作者侧确定性 Judge 的 Golden 正例，不是 Codex/OpenCode 难度分数。四题均未执行无辅助模型 rollout；Codex Harbor 安装和 OpenCode 渠道阻塞仍按上文记录，难度保持 `uncalibrated`。C013 的 `solution/Golden/calendar_source_evidence.xlsx` 仅为通用资产校验所需的输入证据副本，不是额外 solver 输出。

### 第二批滚动制题结果（2026-08-29）

| 候选 | Harbor 包 | 作者假设 | Golden/Judge | 负例与渲染 | 当前状态 |
|---|---|---|---|---|---|
| XLSX-C007 | `tasks/xlsx/officebench-1-5-budget-growth-forecast` | L2 | Golden、fresh replay、5 次重复均 `1.000000`；公式变异重算通过 | 空/损坏/未修改/硬编码/错误公式/非目标篡改；Calc 1 页 | `pending_human_review` |
| XLSX-C020 | `tasks/xlsx/officebench-3-8-revenue-filter-report` | L3 | Golden、fresh replay、5 次重复均 `1.000000` | 空/缺报告/损坏/未筛选/遗漏/报告冲突；Calc+Writer 各 1 页 | `pending_human_review` |
| XLSX-C085 | `tasks/xlsx/agentcompany-budget-variance-analysis` | L2 | Golden、fresh replay、5 次重复均 `1.000000`；420 键独立联结 | 空/损坏/未分析/遗漏/误标/来源值/计算值错误；Calc 横向 1 页 | `pending_human_review` |
| XLSX-C093 | `tasks/xlsx/cowork-bench-financial-audit-reconciliation` | L3 | Golden、fresh replay、5 次重复均 `1.000000`；三输入独立重算 | 空/缺 ICS/损坏/错分/遗漏/总额/原因/邮件冲突；Calc 2 页+Writer 1 页 | `pending_human_review` |

四题通过 `validate-task`、`validate-task-assets`；新增后全库 `validate-public-tasks tasks` 通过 65 个任务，`git diff --check` 通过，制题期间产生的 `__pycache__`/`.pyc` 已清理。以上 `1.000000` 仍只代表作者侧 Golden 正例，不能用于难度标定。C093 将实际邮件/日历副作用离线化为 EML/ICS 草稿，需人工确认适配充分性。

### 第三批滚动制题结果（2026-08-30）

| 候选 | Harbor 包 | 作者假设 | Golden/Judge | 负例与渲染 | 当前状态 |
|---|---|---:|---|---|---|
| XLSX-C002 | `tasks/xlsx/officebench-1-5-shopping-quantity` | L1 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.000000`；未修改=`0.650000`；错改 cabbage=`0.450000`；Calc 1 页 | `pending_human_review` |
| XLSX-C003 | `tasks/xlsx/officebench-1-7-delete-student` | L1 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.000000`；未修改=`0.250000`；误删 Bob=`0.450000`；Calc 1 页 | `pending_human_review` |
| XLSX-C005 | `tasks/xlsx/officebench-1-11-score-sort` | L1 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.000000`；未排序=`0.600000`；Judge 独立拒绝错误键/遗漏行；Calc 1 页 | `pending_human_review` |
| XLSX-C011 | `tasks/xlsx/officebench-1-11-score-split` | L2 | 两个 Golden、fresh replay 均 `1.000000` | 空/损坏=`0.000000`；交换两个输出=`0.400000`；缺文件触发 R001；Calc 各 1 页 | `pending_human_review` |

第三批输入均来自 OfficeBench 固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，各题输入哈希、原始 Query、许可和改造边界记录在对应 `adaptation/README.md`。C003 将上游较弱的关键字差分检查加强为独立行序列重算；C005/C011 保留上游 reference 仅作审计证据，Judge 从输入独立计算，不比较 Golden ZIP 字节。C004 暂未制作，因为其上游 evaluator 给出 `80.5`，而输入 20 行的独立平均值为 `80.05`，证据冲突需单独裁决。

以上分数仍只代表作者侧确定性 Judge 正负例验证，不是 Codex/OpenCode 难度分数。四题均未执行无辅助模型 rollout，难度保持 `uncalibrated`，由人工复核决定接受、修改、降级或淘汰。

### 第四批滚动制题结果（2026-08-30）

| 候选 | Harbor 包 | 作者假设 | Golden/Judge | 负例与渲染 | 当前状态 |
|---|---|---:|---|---|---|
| XLSX-C006 | `tasks/xlsx/officebench-1-14-score-column-swap` | L1 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.000000`；未修改=`0.400000`；只换值不换表头=`0.600000`；Calc 1 页 | `pending_human_review` |
| XLSX-C010 | `tasks/xlsx/officebench-1-9-score-nearest-average` | L2 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.000000`；未修改=`0.100000`；追加错误记录=`0.450000`；Calc 1 页 | `pending_human_review` |
| XLSX-C012 | `tasks/xlsx/officebench-1-13-graduation-check` | L2 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.000000`；输出 Yes=`0.850000`；带解释文本=`0.650000`；输入课程表 Calc 1 页 | `pending_human_review` |

第四批输入均来自 OfficeBench 固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，输入哈希、原始 Query、许可、reference 证据和改造边界记录在对应 `adaptation/README.md`。C010 明确按所有 20 条输入记录计算平均值 `80.05`，不沿用 C004 上游 evaluator 的冲突值 `80.5`；C012 的隐藏 `solution/Golden/course_source_evidence.xlsx` 仅为公开资产校验所需的输入证据副本，不是额外 solver 输出。

以上分数仍只代表作者侧确定性 Judge 正负例验证，不是 Codex/OpenCode 难度分数。三题均未执行无辅助模型 rollout，难度保持 `uncalibrated`，由人工复核决定接受、修改、降级或淘汰。

### 后续滚动策略

Codex Harbor Agent 安装和 OpenCode Opus provider 的基础设施阻塞尚未改变；已完成的 46 道候选均等待项目负责人结合真实 rollout 完成人工复核。本轮按用户要求收口到 60 道，剩余 `ready` 候选和 `conditional` 候选暂不制作。C004 的上游 `80.5` 与输入 20 条记录的独立平均值冲突已在第九批按可审计输入解决为 `80.05`，原始冲突仍保留在候选清单作为来源证据。

### 第五批滚动制题结果（2026-08-30）

| 候选 | Harbor 包 | 作者假设 | Golden/Judge | 负例与渲染 | 当前状态 |
|---|---|---:|---|---|---|
| XLSX-C015 | `tasks/xlsx/officebench-2-15-calendar-to-xlsx` | L3 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.100000`；错误联系人=`0.750000`；Calc 横向 1 页 | `pending_human_review` |
| XLSX-C017 | `tasks/xlsx/officebench-2-24-score-average` | L2 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.050000`/`0.200000`；错误平均值=`0.700000`；Calc 1 页 | `pending_human_review` |
| XLSX-C018 | `tasks/xlsx/officebench-3-1-task-pdf-email` | L3 | Golden、fresh replay 均 `1.000000` | 空/损坏=`0.050000`；缺附件=`0.650000`；PDF 1 页+MIME 附件一致 | `pending_human_review` |

第五批输入均来自 OfficeBench 固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`。C015 从日历事件独立重建三行可编辑议程表；C017 从三条成绩独立计算平均值 `79.0`，不把上游只查关键词的 evaluator 当作完整真值；C018 用 LibreOffice 离线转换完整 12 行任务表，并以 MIME 解析和附件字节一致性替代上游的弱关键词检查。C021 因输入重复姓名、缺失 Tom 且 evaluator 排除实际有进步的 Frank，证据冲突暂缓制作。三题均保持 `uncalibrated`，等待无辅助 rollout 和人工复核。

### 第六批滚动制题结果（2026-08-30）

| 候选 | Harbor 包 | 作者假设 | Golden/Judge | 负例与渲染 | 当前状态 |
|---|---|---:|---|---|---|
| XLSX-C082 | `tasks/xlsx/agentcompany-respondent-format` | L2 | Golden、fresh replay 均 `1.000000`；输入哈希和 unformatted 保护独立检查 | 空/损坏=`0.000000`；缺 formatted=`0.230000`；错误填色=`0.800000`；未居中=`0.850000`；源表篡改=`0.800000`；LibreOffice PDF | `pending_human_review` |
| XLSX-C014 | `tasks/xlsx/officebench-2-14-score-calendar` | L3 | Golden、fresh replay 均 `1.000000`；前三名由输入稳定重算 | 空/损坏/错误排名/错误时间=`0.100000`；额外学生=`0.700000`；ICS 结构独立解析 | `pending_human_review` |
| XLSX-C022 | `tasks/xlsx/officebench-3-52-agenda-calendar-email` | L3 | Golden、fresh replay 均 `1.000000`；议程时间、联系人和跨文件内容由输入独立重算 | 空/损坏/缺 `Charlie.eml`=`0.500000`；错误时间、议题、收件人、额外文件均低于通过阈值；ICS/EML 独立解析 | `pending_human_review` |

C082 来源为 TheAgentCompany 固定 commit `98b68ef82a47690c316f42fddb05baafaab56851`，输入 `respondent_counts.xlsx` SHA256 `be71871d79156a2106b085cf29a20e08ee39bf305c82022003f9533d33ffc043`；仓库根部 MIT，未发现该二进制的单独条款，LICENSE 和风险记录保留在题内 `adaptation/`，最终再发布权利待人工复核。任务从上游 OwnCloud 下载/上传改为本地单工作簿交付，Judge 重新比较 21×11 内容、18 个合并区域、指定 18 个表头填充和所有非空单元格居中。

C014 来源为 OfficeBench 固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，输入 `score.xlsx` SHA256 `a5c8026dc09f38411c6b8783dad74f7f66f08387eeb21b9d8eb7f28a90dd5b6e`；从 20 行记录独立重算 `midterm1` 前三名 Tina=100、Rachel=100、Sam=99，将原日历副作用离线化为三个 UTF-8 VEVENT 文件。上游未提供会议结束时间、地点、邮箱或时区，适配未增加这些要求。

C022 来源为 OfficeBench 固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，输入 `meeting_agenda.xlsx` SHA256 `ffd3161364fde910fbd5491cc1dd19ba346134231e60f217ad668a874af80b0e`；从三条议程独立重建 Bob 的单个日历事件和 Linda、Davide、Charlie 三封纯文本 MIME 提醒，将原日历/邮件副作用离线化为四个本地文件。上游未提供时区、邮箱或附件要求，适配未增加这些事实；缺 `Charlie.eml` 的真实缺失文件负例评分为 `0.500000`。三题均保持 `uncalibrated`，尚未执行无辅助 rollout，等待人工复核。

### 第七批滚动制题结果（2026-08-31）

| 候选 | Harbor 包 | 作者假设 | Golden/Judge | 负例与渲染 | 当前状态 |
|---|---|---:|---|---|---|
| XLSX-C095 | `tasks/xlsx/cowork-bench-faculty-workload-review` | L3 | Golden、fresh replay 均 `1.000000`；33 名教师、23 名超负荷教师和 7 个学科汇总从快照独立重算 | 空/缺 PPTX/损坏=`0.000000`；单教师数值错误=`0.660000`；PPT 总数不一致=`0.640000`；额外文件=`0.880000` 且 hurdle 不通过；Calc/Impress 4 个 PPT 页、Excel 4/3 页 | `pending_human_review` |

C095 来源为 Cowork-Bench 固定 commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`，数据库快照 `db/init.sql.gz` SHA256 `bd9c45b5ec9072b9f301b7df5b3e9ac65540ffcd53485329ebea44b0c3710da4`；保留原始教师目录、政策 PDF 和规范 JSON，从同一 Canvas SQL 快照投影 23 门课程、41 条教师授课关系、32,605 条学生选课记录和 206 条作业记录。原 Canvas API 和 Google Sheet 副作用离线化为四个 CSV 和 `Faculty_Workload_Analysis.xlsx`；Judge 从投影逐行重算，不读取上游 Ground Truth。上游 Golden 对 85.25/145.25 的一位小数使用半偶舍入，适配按 Excel ROUND 的半舍入规则得到 85.3/145.3，差异和部门按授课关系累加语义均记录在题内 `adaptation/README.md`，待人工裁决。任务保持 `uncalibrated`，尚未执行无辅助 rollout。

### 第八批滚动制题结果（2026-08-31）

| 候选 | Harbor 包 | 作者假设 | Golden/Judge | 负例与重算验证 | 当前状态 |
|---|---|---:|---|---|---|
| XLSX-C031 | `tasks/xlsx/officeval-093-editable-overlapping-pentagons` | L1 | Golden、fresh replay、Judge 均 `1.000000` | OOXML 可编辑双五边形、输入保留和 LibreOffice 打开检查 | `pending_human_review` |
| XLSX-C016 | `tasks/xlsx/officebench-2-19-student-teaching-report` | L3 | Golden、fresh replay、Judge 均 `1.000000` | Word 原生报告、12 名学生统计、班级平均分 79.35 和 A/B/C/D/F 分布 5/6/3/6/0 独立重算 | `pending_human_review` |
| XLSX-C045 | `tasks/xlsx/spreadsheetbench2-12-01-financial-model-completion` | L3 | Golden、fresh replay、Judge 均 `1.000000` | Assumption、Working Capital、P&L 目标公式独立构造，并用 LibreOffice 重算；非目标单元格容差检查 | `pending_human_review` |
| XLSX-C060 | `tasks/xlsx/spreadsheet-rl-loan-fifo` | L3 | Golden、fresh replay、Judge 均 `1.000000` | 从 Transactions 独立 FIFO 重算 7 笔 SELL，期望 Avg Cost 为 94.0、94.375、94.8、95.0、96.0、96.14、96.7 | `pending_human_review` |
| XLSX-C061 | `tasks/xlsx/spreadsheet-rl-salary-forecast` | L3 | Golden、fresh replay、Judge 均 `1.000000` | 从增幅表、级别/步级和日期独立重算第 8–15 行 M/O/Q 公式，并用 LibreOffice 重算 | `pending_human_review` |

第八批新增 1 道 L1、4 道 L3，全部保持 `uncalibrated` 和 `pending_human_review`，尚未执行无辅助模型 rollout。C023 未制作并由 C045 替换：其“推荐四门课”任务缺少可审计的课程目录，第四门推荐无法形成输入到真值的闭环。第八批完成后，xlsx 目录共有 41 道题（现有 14 道基线加 27 道新候选），全库公开端到端任务静态校验为 `84/84`；上述分数仅代表作者侧确定性 Judge 的正例验证，不是难度标定结果。

## Codex + GPT-5.6-sol 冻结 Rollout（2026-08-30）

已对上面 8 个已完成候选执行 `xlsx-codex-gpt56sol-8x3-20260830`：每题 3 次，共 24 次；Codex CLI `0.150.1`、模型 `gpt-5.6-sol`、`skills=[]`、无 Memory/few-shot/制题材料；每次为全新 Harbor Docker trial，隐藏侧 Judge 在模型执行后运行。运行结果保存在 `.coworkevo/runs/xlsx-codex-gpt56sol-8x3-20260830/`，汇总 `n_completed_trials=24`、`n_errored_trials=0`、`agent_execution` 覆盖率 `24/24`；整体 Judge 均值 `0.758333`，总输入 `4,602,450` tokens、缓存 `4,018,176`、输出 `125,838`、成本约 `$8.755237`。这些是首次真实模型结果，但仍属于 `uncalibrated` 证据，不自动替代人工接收、修改、降级或淘汰决定。

| Harbor 包 | 作者假设 | n | Judge 均值 | 标准差 | `score>=0.7` | 主要观察 |
|---|---:|---:|---:|---:|---:|---|
| `officebench-1-4-score-update` | L1 | 3 | 1.000000 | 0.000000 | 3/3 | 稳定通过 |
| `officebench-1-6-class-schedule` | L2 | 3 | 0.888889 | 0.078568 | 3/3 | 两次 0.833333，一次 1.000000 |
| `officebench-2-13-section-calendar` | L3 | 3 | 0.611111 | 0.157135 | 1/3 | ICS 边界、标题/描述和时间格式是主要扣分点 |
| `cowork-bench-supplier-contract-compliance` | L4 | 3 | 0.366667 | 0.047140 | 0/3 | 评分/整改/逐供应商报告/通知/日历/跨文件一致性均有缺陷 |
| `officebench-1-5-budget-growth-forecast` | L2 | 3 | 0.950000 | 0.000000 | 3/3 | 稳定通过，公式相关要求基本满足 |
| `officebench-3-8-revenue-filter-report` | L3 | 3 | 0.950000 | 0.000000 | 3/3 | 稳定通过 |
| `agentcompany-budget-variance-analysis` | L2 | 3 | 0.700000 | 0.000000 | 3/3 | 固定遗漏一个关键计算/格式判据，恰过阈值 |
| `cowork-bench-financial-audit-reconciliation` | L3 | 3 | 0.600000 | 0.108012 | 1/3 | 差异原因及 DOCX/EML/ICS 一致性交付不稳定 |

三次分数依次保留在各 trial 的 `artifacts/logs/artifacts/judge-result.json`；未发生安装、鉴权、模型不可用、Agent timeout 或 verifier exception。低分不是 Harbor 的 `Mean 0.000` 占位异常，而是模型已实际执行后的确定性 Judge 结果。分段日历的一次 0.833333、两次 0.500000；财务对账为 0.500000、0.550000、0.750000；供应商合规为 0.300000、0.400000、0.400000，说明高复杂度跨应用任务仍有明显执行波动。

本次基础设施修复使用外部 Docker volume `cowork-codex-0-150-1` 预装 Node `v22.23.2`、Codex `0.150.1` 和 ripgrep，并通过 `.coworkevo/configs/codex-runtime-volume.compose.yaml` 只读挂载；Codex job 显式注入 `${ZHICLOUD_API_KEY}` 和 `METHOD=script`，并在临时 `CODEX_HOME` 写入 ZhiCloud Responses 配置。安装预检和单题真实 smoke 均通过后才启动 24 次矩阵。该 volume/config 属于本机 rollout 基础设施，不含密钥，不应作为题目可见内容或作者 Skill 提交。

### 第九批滚动制题结果（2026-08-31）

| 难度 | 新增任务数 | Harbor 包 | Golden/fresh replay | 负例验证 | 当前状态 |
|---|---:|---|---|---|---|
| L1 | 3 | `officebench-1-9-score-average`、`spreadsheet-rl-ups-power`、`spreadsheet-rl-clean-email-body` | 3/3 均 `1.000000` | 空/损坏均失败；未修改输入 `0.20–0.40` | `pending_human_review` |
| L2 | 8 | `officebench-1-6-budget-rules`、`sb2-template-06-19-revenue-cagr`、`sb2-template-14-09-concession-revenue`、`sb2-visualization-1437004-activity-dashboard`、`sb2-visualization-129-mars-valuation`、`spreadsheet-rl-cross-midnight-hours`、`spreadsheet-rl-inventory-eom`、`spreadsheet-rl-square-quickbooks-master` | 8/8 均 `1.000000` | 空/损坏均失败；未修改输入 `0.15–0.55` | `pending_human_review` |
| L3 | 2 | `officebench-3-6-class-training-email`、`sb2-financial-model-13-02-logistics` | 2/2 均 `1.000000` | 培训包缺文件失败；物流模型未修改输入 `0.40` | `pending_human_review` |
| L4 | 6 | `sb2-debugging-01-03-lbo-repair`、`sb2-debugging-02-03-apc-merger-repair`、`sb2-debugging-03-03-lbo-repair`、`sb2-financial-model-03-02-bank-valuation`、`sb2-financial-model-05-01-aif-fund`、`sb2-financial-model-07-02-real-estate` | 6/6 均 `1.000000` | 空/损坏均失败；未修改输入 `0.15–0.40` | `pending_human_review` |

第九批使用 OfficeBench、Spreadsheet-RL 和 SpreadsheetBench-2 固定公开 revision；每题均保留输入 SHA256、来源 URI、许可和改造边界。批量 Judge 已修正追加结果工作表、工作表重排和 OOXML 数组公式的比较逻辑，参考解法支持 `OUTPUT_DIR` fresh replay。`officebench-3-6-class-training-email` 的隐藏 Golden 额外保留 `training_schedule_reference.xlsx` 作为资产校验所需的输入证据副本，不属于声明输出。批次完成后 `validate-public-tasks tasks` 通过 `103/103`；全库共 60 道，实际分布为 L1=9、L2=18、L3=25、L4=8。所有新增题仍为 `uncalibrated`，未执行无辅助模型 rollout，等待人工复核。

### 第十批补题与最终配额（2026-08-31）

为回应 Sonnet 5 报告中 Basic 覆盖不足的问题，继续依据 OfficeBench 固定 revision `b978b808667c32b52ce19a67ce1def1de9ae02b7` 制作 6 道单工作簿 L1 题：`officebench-1-6-attendance`、`officebench-1-10-highest-salary-duplicate`、`officebench-1-10-lowest-budget-add-highest`、`officebench-1-9-total-salary`、`officebench-1-9-total-budget`、`officebench-1-9-total-shopping-weight`。每题保留原始子任务 JSON、Apache-2.0 许可证、输入 SHA256 和适配边界；参考解法与 Judge 均从输入独立计算。作者侧 fresh replay 和确定性 Judge 均为 `1.000000`，但尚未执行无辅助模型 rollout，难度仍为未校准假设。

综合任务结构、已有多模型/失败原因记录和重复度，对 4 道真实异构、多交付或状态依赖任务记录 L4 结构假设：`cowork-bench-customer-satisfaction-analysis-report`、`cowork-bench-json-data-cleaning-pipeline`、`cowork-bench-moex-sector-rotation-dashboard`、`officeval-082-checkbox-driven-tag-summary`。这不是依据 Sonnet 单次得分的自动重标，4 题均保持 `calibration_status=uncalibrated` 与 `review_status=pending_human_review`。

为保持上一版活动集恰为 60 道并平衡档位，6 道结构重复候选曾移至 `datasets/xlsx-archive/`。本轮按最新边界只归档 `officebench-1-13-graduation-check` 和 `officebench-2-14-score-calendar` 两道非 XLSX 输出题，并恢复已审计的 `officebench-2-15-calendar-to-xlsx` 与 `sb2-template-14-09-concession-revenue`；5 道 Cross Office 题和 1 道 Word 题保持原目录，待学长迁移。当前 `tasks/xlsx` 工作目录的 60 包计数为 **L1=15、L2=17、L3=16、L4=12**，但仅 **54** 道已完成唯一 XLSX 输出合同；6 道待交接包不应纳入新的 Excel rollout 或正式 Release。

## 历史：Claude Code + Sonnet 5 最终 60 题 Rollout（2026-09-01）

该次 rollout 在上述重分类、归档、改造和候选恢复之前完成，共 60 题 × 3 次、180 个 trials。它保留为历史基线，不能作为 7 道重制题或 2 道恢复题的新版评分，也不能视为 54 道现行 Excel 合同集的统计。运行汇总为 `.coworkevo/runs/xlsx-60-final-complete-summary.json`；提交版题级清单为 [`xlsx-rollout-20260901/manifest.json`](xlsx-rollout-20260901/manifest.json) 和 [`xlsx-rollout-20260901/task-results.csv`](xlsx-rollout-20260901/task-results.csv)。

| 指标 | 结果 |
|---|---:|
| 题目 / trials | 60 / 180 |
| 平均 reward | 0.438649（43.86/100） |
| 满分题（题级均值=1） | 13 |
| Reward ≥ 0.8 题（best-of-3） | 19 |
| 非零分题（题级均值>0） | 47 |
| 零分题（题级均值=0） | 13 |
| 单次通过（reward≥0.7） | 56/180（31.1%） |
| 难度分布 | L1/L2/L3/L4 = 15/16/17/12 |

按难度的 trial 通过率为 L1 80.0%（36/45）、L2 39.6%（19/48）、L3 2.0%（1/51）、L4 0.0%（0/36）。L1→L2 和 L2→L3 分别下降 32.27 和 32.87 分，显示 L2→L3 存在能力断层；L3/L4 之间只差 3.57 分，需优先审计可解性和交付链，而不是简单提高难度。

本次运行报告记录总量 **120,256,921 tokens、$113.73**；题级 summary 本身不含 token 字段，提交版 manifest 已明确标记该成本为 run-level reported accounting。60 个题包仍全部保持 `pending_human_review`，rollout 结果不自动改变 `metadata/task.json` 的难度或发布状态。
