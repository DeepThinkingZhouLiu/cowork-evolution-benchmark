# PPT 60 题建设台账 v0.1

## 台账边界

本文档只记录选题和逐题建设进度，不是题目生成器。每道题仍必须按 `.agents/skills/cowork-pptx-task-authoring/SKILL.md` 单独读入上游 Query、evaluator 和全部输入，然后才能制作 Golden、Rubric 和 Judge。

`design_target` 只是制题阶段的候选档位。任务的 `difficulty_level` 和 `difficulty_tier` 在真实 Agent Rollout 前一律保持 `unassigned`。

## 建设配比

| 范围       | L1 / Easy | L2 / Medium | L3 / Hard | L4 / Expert | 合计 |
|------------|-----------|-------------|-----------|-------------|------|
| 现有候选 | 1         | 4           | 8         | 6           | 19   |
| 新增候选 | 11        | 14          | 10        | 6           | 41   |
| 建设目标 | 12        | 18          | 18        | 12          | 60   |

现有 19 题的这一分布只是用于控制新增 Seed 的结构预估，不写回正式难度字段。如果后续逐题审计或 Rollout 改变了分布，应从已审计的备选 Seed 中替换，不得为凑数修改 Judge 或难度阈值。

## 当前验收结果

- PPT 任务目录、`task.toml`、真实输入包、Golden、Rubric、Judge 和 Oracle 解法为 `60/60`。
- 当前工作树的 `jobs/` 中可直接追溯 `60/60` 道任务的官方 Harbor 0.21 Oracle 成功 trial；全部 Reward 为 1.0、异常数为 0，并实际收集到 `/app/output` 产物。
- 2026-09-04 已对 Claude Code 2.1.251、OpenCode 1.4.14 和 OpenHands SDK 1.44.1 的 180 份 Sonnet 5 答案完成 `70%` 代码分与 `30%` Codex-VLM 分的试校准；179 份正常完成、1 份真实任务失败、0 份评测基础设施失败。试校准按 Panel 平均分排序后得到 12 / 18 / 18 / 12，逐题结果见 `benchmarks/difficulty-labels-v0.5.csv`。
- 旧 19 道已统一迁移到 `environment/input_files/`、`/app/input_files`、标准 Reward、`network_mode = "no-network"` 和单一 `/app/output` 采集协议。
- 旧 19 道的空交付和损坏文件共 38 个负例全部为 0；11 道可直接回交原始 PPTX 的任务均低于通过阈值。`officeval-075` 在这一轮发现并修复了页数不足时 Judge 越界崩溃的问题，修复后负例为 0、Golden 和最终 Harbor 回放为 1.0。
- `L1 / L2 / L3 / L4` 的12/18/18/12已于2026-09-17按既有试校准结果回填到60题的TOML/JSON；状态为 `pilot_single_panel`，多模型、多次重复验证未完成。统一索引见 `benchmarks/difficulty-labels-v0.5.csv`。

## 现有 19 题 Harbor 回放

| #  | 任务目录 | 来源 | 家族 | 最终成功 job / trial | 状态 |
|----|----------|------|------|----------------------|------|
| 01 | `cowork-bench-fy2026-leadership-budget-review` | Cowork-Bench | `audience_adaptation` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `cowork-bench-fy2026-leadership-b__2rWxbVX` | `harbor_replayed` |
| 02 | `cowork-bench-q4-executive-briefing-package` | Cowork-Bench | `presentation_package` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `cowork-bench-q4-executive-briefi__3ZqwRrR` | `harbor_replayed` |
| 03 | `officeval-041-note-driven-deck-audit` | OmegaUse-OfficeVal | `deck_audit` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-041-note-driven-deck-a__B82kw33` | `harbor_replayed` |
| 04 | `officeval-043-business-card-to-editable-bilingual-deck` | OmegaUse-OfficeVal | `material_to_deck` | `ppt-legacy-standard-seven-oracle-rerun-20260829` / `officeval-043-business-card-to-e__PXjaNf5` | `harbor_replayed` |
| 05 | `officeval-046-image-to-editable-dual-research-diagrams` | OmegaUse-OfficeVal | `material_to_deck` | `ppt-legacy-standard-seven-oracle-rerun-20260829` / `officeval-046-image-to-editable__YhxqJBG` | `harbor_replayed` |
| 06 | `officeval-051-image-to-editable-anomaly-detection-flow` | OmegaUse-OfficeVal | `material_to_deck` | `ppt-legacy-standard-seven-oracle-rerun-20260829` / `officeval-051-image-to-editable__X7seMJG` | `harbor_replayed` |
| 07 | `officeval-054-image-to-editable-water-reuse-process` | OmegaUse-OfficeVal | `material_to_deck` | `ppt-legacy-standard-seven-oracle-rerun-20260829` / `officeval-054-image-to-editable__WzsUmru` | `harbor_replayed` |
| 08 | `officeval-055-image-to-editable-outage-flow` | OmegaUse-OfficeVal | `material_to_deck` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-055-image-to-editable__7eNDrtd` | `harbor_replayed` |
| 09 | `officeval-056-image-to-editable-complaint-closure-flow` | OmegaUse-OfficeVal | `material_to_deck` | `ppt-legacy-standard-seven-oracle-rerun-20260829` / `officeval-056-image-to-editable__QTAAjLd` | `harbor_replayed` |
| 10 | `officeval-060-single-slide-image-layout-repair` | OmegaUse-OfficeVal | `visual_editing` | `ppt-legacy-standard-seven-oracle-rerun-20260829` / `officeval-060-single-slide-image__3TPtvin` | `harbor_replayed` |
| 11 | `officeval-063-multi-page-content-editing` | OmegaUse-OfficeVal | `content_editing` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-063-multi-page-content__T85qQhY` | `harbor_replayed` |
| 12 | `officeval-064-ordered-photo-layout-completion` | OmegaUse-OfficeVal | `template_based_generation` | `ppt-legacy-standard-seven-oracle-rerun-20260829` / `officeval-064-ordered-photo-layo__j4bk2ah` | `harbor_replayed` |
| 13 | `officeval-066-excel-to-editable-data-deck` | OmegaUse-OfficeVal | `data_to_deck` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-066-excel-to-editable__m5SrESU` | `harbor_replayed` |
| 14 | `officeval-068-editable-deck-consolidation` | OmegaUse-OfficeVal | `deck_consolidation` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-068-editable-deck-cons__QuPoc57` | `harbor_replayed` |
| 15 | `officeval-069-reference-guided-visual-editing` | OmegaUse-OfficeVal | `visual_editing` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-069-reference-guided-v__gzxRRQ4` | `harbor_replayed` |
| 16 | `officeval-070-single-slide-layout-polishing` | OmegaUse-OfficeVal | `deck_polishing` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-070-single-slide-layou__akwowmj` | `harbor_replayed` |
| 17 | `officeval-074-full-deck-localization` | OmegaUse-OfficeVal | `localization` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-074-full-deck-localiza__QiGMUE3` | `harbor_replayed` |
| 18 | `officeval-075-template-style-transfer` | OmegaUse-OfficeVal | `template_based_generation` | `officeval075-oracle-final-20260829` / `officeval-075-template-style-tra__BusThoA` | `harbor_replayed` |
| 19 | `officeval-076-outline-to-animated-sales-plan` | OmegaUse-OfficeVal | `outline_to_deck` | `ppt-legacy-migrated-twelve-oracle-rerun-20260829` / `officeval-076-outline-to-animate__B9AL8Ph` | `harbor_replayed` |

## 新增 41 题候选

| #  | 上游与题号                       | 任务摘要                              | 家族                        | 来源方式  | 目标 | 当前状态             |
|----|------------------------------------|---------------------------------------|-----------------------------|-----------|------|----------------------|
| 01 | PPTArena Case 3                    | 在第 5 页底部添加精确来源                 | content_editing             | direct    | L1   | `harbor_replayed`    |
| 02 | PPTArena Case 4                    | 第 8 页非主标题文字统一增加 2 pt           | content_editing             | direct    | L1   | `harbor_replayed`    |
| 03 | PPTArena Case 14                   | 加粗除封面外所有幻灯片标题                | content_editing             | direct    | L1   | harbor_replayed      |
| 04 | PPTArena Case 15                   | 加粗整份演示文稿的所有标题                  | content_editing             | direct    | L1   | harbor_replayed      |
| 05 | PPTArena Case 21                   | 将所有幻灯片标题规范为 Title Case          | content_editing             | direct    | L1   | `harbor_replayed`    |
| 06 | PPTArena Case 23                   | 在末页追加 Thank you 结束页                | content_editing             | direct    | L1   | `harbor_replayed`    |
| 07 | PPTArena Case 29                   | 将第 3 页柱状图改为可编辑饼图              | visual_editing              | direct    | L1   | `harbor_replayed`    |
| 08 | PPTArena Case 37                   | 为所有幻灯片设置 Fade 转场                  | visual_editing              | direct    | L1   | `harbor_replayed`    |
| 09 | PPTArena Case 41                   | 从图片取色并修改幻灯片背景                   | visual_editing              | direct    | L1   | `harbor_replayed`    |
| 10 | PPTArena Case 59                   | 将文本框中的两个 URL 设为可点击链接          | content_editing             | direct    | L1   | `harbor_replayed`    |
| 11 | PPTArena Case 82                   | 为每页内容添加细深灰色边框                 | visual_editing              | direct    | L1   | `harbor_replayed`    |
| 12 | PPTArena Case 16                   | 矫正并等大对齐三张照片与标签                  | visual_editing              | direct    | L2   | `harbor_replayed`    |
| 13 | PPTArena Case 25                   | 根据现有配色重做表格表头和交替行               | deck_polishing              | direct    | L2   | `harbor_replayed`    |
| 14 | PPTArena Case 27                   | 统一图片占位符尺寸、水平对齐并匹配标签          | visual_editing              | direct    | L2   | `harbor_replayed`    |
| 15 | PPTArena Case 28                   | 重排页面顺序并同步修复议程                    | deck_audit                  | direct    | L2   | `harbor_replayed`    |
| 16 | PPTArena Case 30                   | 将文本数据改为带标记折线图与自定义图例            | data_to_deck                | direct    | L2   | `harbor_replayed`    |
| 17 | PPTArena Case 31                   | 修复越出页面的文本框并保留全部文字                | visual_editing              | direct    | L2   | `harbor_replayed`    |
| 18 | PPTArena Case 36                   | 将第 2 页文本分配到对应页讲者备注后删页            | presentation_package        | direct    | L2   | `harbor_replayed`    |
| 19 | PPTArena Case 44                   | 将旗帜与地区说明匹配并删除重复旗帜               | content_editing             | direct    | L2   | `harbor_replayed`    |
| 20 | PPTArena Case 47                   | 将折线图改为簇状柱形图                      | visual_editing              | direct    | L2   | `harbor_replayed`    |
| 21 | PPTArena Case 50                   | 识别多种日期表达并统一成 Month Day, Year     | content_editing             | direct    | L2   | `harbor_replayed`    |
| 22 | PPTArena Case 53                   | 把文本数据准确写入可编辑表格并删除原文本框          | data_to_deck                | direct    | L2   | `harbor_replayed`    |
| 23 | PPTArena Case 61                   | 按 Score 降序重排整张表格                    | content_editing             | direct    | L2   | `harbor_replayed`    |
| 24 | PPTArena Case 64                   | 将图表系列改为 Okabe-Ito 色盲友好配色          | visual_editing              | direct    | L2   | `harbor_replayed`    |
| 25 | PPTArena Case 71                   | 跨全部表格识别负数并设置红色加粗                  | content_editing             | direct    | L2   | `harbor_replayed`    |
| 26 | PPTArena Case 1                    | 哈萨克语译英语，完整保留法语和格式                 | localization                | direct    | L3   | `harbor_replayed`    |
| 27 | PPTArena Case 67                   | 修复唯一标题、阅读顺序和主题字体可访问性             | deck_audit                  | direct    | L3   | `harbor_replayed`    |
| 28 | PPTArena Case 68                   | 将混乱流程重构为三泳道可编辑规范图               | visual_editing              | direct    | L3   | `harbor_replayed`    |
| 29 | PPTArena Case 74                   | 跨页合并绩效数据、删页、统一标题并对齐图片           | deck_consolidation          | direct    | L3   | `harbor_replayed`    |
| 30 | PPTArena Case 83                   | 从页脚去重提取来源，排序并新增参考文献页             | deck_consolidation          | direct    | L3   | `harbor_replayed`    |
| 31 | PPTArena Case 93                   | 主题、图表、图片流、进度条与参考文献的级联修订         | deck_polishing              | direct    | L3   | `harbor_replayed`    |
| 32 | AutomationBench 4040               | 从月结清单生成财务结账就绪简报                 | data_to_deck                | composite | L3   | `harbor_replayed`    |
| 33 | AutomationBench 1115               | 整合追踪表、分析师邮件和销售 Slack 成竞争简报     | material_to_deck            | composite | L3   | `harbor_replayed`    |
| 34 | OfficeBench 3-2 / subtask 0        | 提取 NVIDIA 财报并生成董事会财务摘要演示文稿     | material_to_deck            | composite | L3   | `harbor_replayed`    |
| 35 | OfficeBench 3-5 / subtask 0        | 从两份真实样式发票提取支付趋势并生成评审演示文稿       | material_to_deck            | composite | L3   | `harbor_replayed`    |
| 36 | AutomationBench 4062               | 处理政策例外、税率变更和分州计算的销售税简报          | data_to_deck                | composite | L4   | `harbor_replayed`    |
| 37 | AutomationBench 4044               | 根据政策邮件对发票、订单和收货记录三单匹配并汇报        | data_to_deck                | composite | L4   | `harbor_replayed`    |
| 38 | TheAgentCompany internal tooling   | 跨 RocketChat、GitLab、OwnCloud 与 Plane 制作入职工具演示文稿 | presentation_package        | direct    | L4   | `harbor_replayed`    |
| 39 | TheAgentCompany GitLab progress    | 遍历多个 GitLab 项目的议题和合并请求形成进展演示文稿  | material_to_deck            | direct    | L4   | `harbor_replayed`    |
| 40 | TheAgentCompany customer meeting   | 跨项目与协作系统筛选客户可用 Benchmark 并制作会前演示文稿 | audience_adaptation         | direct    | L4   | `harbor_replayed`    |
| 41 | TheAgentCompany stock analysis     | 清洗股价 CSV，完成多组分析图表并交付演示文稿与讲稿     | presentation_package        | direct    | L4   | `harbor_replayed`    |

## 来源锁定

| 来源             | 锁定版本                                         | 许可与用法                                                        |
|------------------|--------------------------------------------------|-------------------------------------------------------------------|
| PPTArena         | Git `6f26a63e`；HF `20fb8902`                 | MIT；可使用原始 PPTX、人工 Ground Truth 和指令，逐题记哈希  |
| AutomationBench  | Git `4a8e1061`                                   | MIT；使用公开的合成业务状态与断言，诚实标记 synthetic            |
| OfficeBench      | Git `b978b808`                                   | Apache-2.0；保留原任务和 PDF 输入，新增 PPTX 交付为明确适配     |
| TheAgentCompany  | Git `98b68ef8`                                   | MIT；保留原任务、输入和 evaluator，静态化外部应用状态后才发布 |

OSWorld 2.0、WorkArena、AppWorld 和 ClawMark 本轮只用于 L3/L4 结构设计参考，不复制 gated、加密、凭据绑定或非商用限制的 Query、输入、答案或隐藏状态。

## 状态流程

```text
selected
-> source_audited
-> golden_ready
-> judge_ready
-> harbor_replayed
-> rollout_calibrated
-> promoted
```

只有 `promoted` 状态的题目才能进入论文 Benchmark Release。
