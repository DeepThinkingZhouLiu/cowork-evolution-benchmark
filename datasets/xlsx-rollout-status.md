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

## OpenCode 阻塞证据（2026-08-28）

- ZhiCloud `GET /v1/models` 只返回 `gpt-5.5`、`gpt-5.6-sol`、`gpt-5.6-terra`。
- 使用所提供 key 对 `claude-opus-4-8`、`claude-opus-4.8`、`opus-4.8` 和 `claude-opus-4-6` 发起最小 Anthropic Messages 请求，均返回 HTTP 503：当前“GPT 特惠渠道”分组无可用 Claude distributor。
- OpenCode 配置已使用 `{env:ZHICLOUD_API_KEY}`，仓库中未保存密钥。获得带 Claude/Opus 渠道的 key 或调整账户分组后，才能执行真实 Rollout。

## 运行记录要求

每次运行记录任务 commit、Harness/CLI 版本、模型 ID、`skills=[]`、Prompt digest、预算、随机种子、产物、逐条 Rubric、总分、成功与否、失败类型、步骤数、Token 和成本。未实际运行的格子保持空值，禁止填充推测分数。
