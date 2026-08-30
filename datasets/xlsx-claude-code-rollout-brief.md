# CoworkEvoBench Excel 8 题：Claude Code + Claude Opus 4.8 Rollout 说明

这份说明用于让本机 Claude Code CLI 执行已经完成的 8 个 Excel/Cross-Office 任务，获得与 2026-08-30 Codex rollout 可比较的真实难度证据。

## 1. 目标和边界

- Harness：本机 Claude Code CLI，版本运行时记录。
- 模型：`claude-opus-4-8`。不要改成 `claude-opus-4.8`、`opus-4.8` 或其他模型。
- Provider：ZhiCloud Anthropic-compatible API。
- 重复：每道题 3 次，必须是相互独立的新工作目录，共 24 次 trial。
- 评测协议：`skills=[]`，不挂载 Memory、few-shot、作者制题 Skill、Golden、Rubric、Judge 或其他隐藏辅助材料。
- 允许读取的求解材料只有该题的 `instruction.md`、`task.toml` 中公开运行字段和 `data/input_files/` 输入文件。
- `solution/`、`tests/`、`rubric.json`、`adaptation/README.md`、`metadata/`、其他 trial 产物和 Codex 结果均不得提供给 Claude。
- 外部副作用保持离线：邮件和日历任务只交付 `.eml`/`.ics` 草稿，不发送邮件、不创建真实日历事件。

Codex + GPT-5.6-sol 挂载 Excel 制题 Skill 只用于作者侧制作和验证任务包，不属于这次被测模型的能力输入。不要把该 Skill 复制到 Claude 的工作目录。

## 2. API 配置

在运行 Claude Code 的同一个 shell 会话中设置变量。不要把真实 API key 写入仓库、脚本、日志或本说明；此前在聊天和命令行中暴露过的 key 应先撤销并轮换。

```bash
export ANTHROPIC_BASE_URL="https://api.zcloudapi.com/v1"
export ANTHROPIC_API_KEY="<rotated-key-from-your-shell>"
claude --version
claude -p --model claude-opus-4-8 "Reply with OK only." 
```

最小 smoke 只用于确认 CLI、鉴权和模型路由可用，不计入 24 次正式 trial。若 smoke 失败，记录 HTTP/provider 错误并停止，不要把失败记为题目得分。

可直接粘贴给 Claude Code 的启动消息：

```text
请按 datasets/xlsx-claude-code-rollout-brief.md 执行 Claude Code + claude-opus-4-8 的 8 题 rollout。先完成一次不计分的 API smoke，然后严格按说明创建隔离工作区，依次对 8 道题各运行 3 个全新 trial。求解阶段只能看到对应 instruction.md 和 data/input_files；不要读取 solution、tests、rubric、adaptation、metadata、其他 trial 或 Codex 结果。每次只把交付物写入该 trial 的 output/，退出后保存日志和元数据。Judge 必须在 Claude 退出后的评测侧运行，不能把 Judge 实现或分数提示给求解过程。最后生成完整 summary.json，并报告每题 3 次 reward、失败分类和待人工复核项；不要把任何题自动标记为 accepted/released。
```

## 3. 8 道任务

所有路径相对于仓库根目录 `/data/chenxingyan/cowork-evolution-benchmark`。每次 trial 将允许读取的输入复制到一个全新的临时工作目录，并把输出放在该目录的 `output/` 下；任务说明中的 `/app/output` 在本机运行时等价映射到这个 `output/`。

| task | 作者初始假设 | 期望输出 |
|---|---:|---|
| `tasks/xlsx/officebench-1-4-score-update` | L1 | `score.xlsx` |
| `tasks/xlsx/officebench-1-6-class-schedule` | L2 | `class_member_schedule.xlsx` |
| `tasks/xlsx/officebench-2-13-section-calendar` | L3 | `Alice.ics`, `David.ics`, `Grace.ics`, `Jack.ics` |
| `tasks/xlsx/cowork-bench-supplier-contract-compliance` | L4 | `compliance_assessment.xlsx`, `corrective_actions.xlsx`, `audit_findings.docx`, `audit_summary.docx`, `audit_notifications.eml`, `supplier_review_meetings.ics` |
| `tasks/xlsx/officebench-1-5-budget-growth-forecast` | L2 | `company_budget.xlsx` |
| `tasks/xlsx/officebench-3-8-revenue-filter-report` | L3 | `new.xlsx`, `report.docx` |
| `tasks/xlsx/agentcompany-budget-variance-analysis` | L2 | `flagged_variances.xlsx` |
| `tasks/xlsx/cowork-bench-financial-audit-reconciliation` | L3 | `Audit_Reconciliation_Report.xlsx`, `Audit_Report.docx`, `audit_notification.eml`, `audit_review_meeting.ics` |

求解时只向 Claude 提供对应题目的 `instruction.md` 内容和该题输入文件。不要提前告诉 Claude 任何分数、失败模式、Golden 内容或验收标准之外的解法提示。

## 4. 单次 trial 流程

1. 创建全新的临时目录，例如 `.coworkevo/local-trials/xlsx-claude-opus48-8x3-<date>/<task>__run-01/`。
2. 只复制该题的 `instruction.md` 和 `data/input_files/` 到求解工作区；将输入挂载为只读，将空 `output/` 挂载为可写。
3. 在该工作区启动 Claude Code CLI，使用 `--model claude-opus-4-8`，将任务说明作为用户任务提交。Claude 可以使用本机已有的 Office 文件处理工具，但不得访问仓库中的隐藏评测文件。
4. Claude 结束后，确认输出文件只位于该 trial 的 `output/`，保存 CLI 原始日志、退出码、开始/结束时间和版本信息。
5. 将 trial 输出交给仓库 Judge；Judge 在 Claude 退出后、隔离的评测侧执行，不把 Judge 输出回传给 Claude。
6. 记录每条 Rubric、归一化 reward、`passed`、失败类型和产物 SHA256。模型没有完成任务时使用 `agent_failure`；CLI/API、安装、超时或 Judge 异常必须分别标记，不能伪造模型分数。

建议的产物目录：

```text
.coworkevo/runs/xlsx-claude-opus48-8x3-YYYYMMDD/
├── config.json
├── summary.json
└── <task-id>__<run-id>/
    ├── agent/              # Claude CLI 日志和版本信息
    ├── output/             # Claude 的实际交付物
    └── evaluator/          # judge-result.json、reward.txt、逐条证据
```

评测侧可以使用题包自带的标准入口（把第一个参数替换为该 trial 的 `output/`）：

```bash
ARTIFACT_DIR="<trial>/output" \
JUDGE_RESULT_PATH="<trial>/evaluator/judge-result.json" \
REWARD_PATH="<trial>/evaluator/reward.txt" \
bash "tasks/xlsx/<task-id>/tests/test.sh" "$ARTIFACT_DIR"
```

不要在 Claude 的工作区执行或展示 `tests/test.sh`、`tests/judge.py`、`rubric.json` 或 `solution/solve.*`；上述命令只在求解完成后的隐藏评测侧运行。

## 5. 汇总格式

`summary.json` 至少记录：

```json
{
  "harness": "claude-code",
  "harness_version": "<claude --version>",
  "model": "claude-opus-4-8",
  "provider_base_url": "https://api.zcloudapi.com/v1",
  "skills": [],
  "n_tasks": 8,
  "n_runs_per_task": 3,
  "n_completed_trials": 0,
  "n_errored_trials": 0,
  "task_results": [],
  "failure_taxonomy": {
    "agent_failure": 0,
    "provider_or_auth": 0,
    "setup_or_timeout": 0,
    "verifier_exception": 0
  }
}
```

每道题至少汇总：3 次 reward、均值、标准差、`score >= 0.7` 的次数、是否每次满足核心门槛、主要失败 Rubric 和人工备注。不要把作者侧 Golden 的 `1.0` 当成模型成绩。

## 6. 结果解释和验收

这次 rollout 的用途是校准难度、发现任务/适配问题和支持人工决策，不是自动发布。即使某题 3/3 达到 `0.7`，也仍需人工确认输入解释、离线化副作用、输出专业性和 Rubric 是否合理；低分或高波动题应先修改或降级，再重新跑回归。

请特别检查此前 Codex 暴露的风险区域：Section Calendar 的 VCALENDAR 边界、标题/描述/时间；Supplier Contract 的评分、整改、逐供应商报告、邮件/日历及跨文件一致性；Financial Audit 的差异原因和 DOCX/EML/ICS 一致性；Budget Variance 的关键计算和格式列。

完成 24 次 trial 和人工复核前，不要批量制作剩余候选题，也不要把这 8 题标记为最终 accepted/released。
