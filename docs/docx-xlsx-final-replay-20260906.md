# DOCX/XLSX Harbor Oracle 冻结回放

## 最终结果

本次使用 Harbor 0.21.0、Docker 环境和 Oracle Agent，对 DOCX 60 道与 XLSX 60 道分别执行两轮全量回放。四轮均从独立空输出目录开始，任务运行和验证阶段均使用 `network_mode = "no-network"`。

| 模块 | 任务数 | 回放轮数 | Trial 数 | Reward 1.0 | 异常 | 两轮任务 checksum 不一致 | OOXML 有效负载不一致 |
|------|-------:|---------:|---------:|-----------:|-----:|--------------------------:|---------------------:|
| DOCX |     60 |        2 |      120 |        120 |    0 |                         0 |                    0 |
| XLSX |     60 |        2 |      120 |        120 |    0 |                         0 |                    0 |
| 合计 |    120 |        4 |      240 |        240 |    0 |                         0 |                    0 |

最终作业：

- DOCX：`docx-60-final-a-20260906`、`docx-60-final-b-20260906`
- XLSX：`xlsx-60-final-a-20260906`、`xlsx-60-final-b-20260906`

DOCX 每轮收集 62 个交付文件，4 道题两轮产物逐字节一致，56 道题的 OOXML 正文、结构和媒体负载一致；XLSX 每轮收集 60 个交付文件，21 道题逐字节一致，39 道题的 OOXML 有效负载一致。非逐字节一致项只包含 ZIP 时间戳或 `docProps/core.xml` 的 created/modified 时间差异。

## 本轮修复

- 为 8 道历史 Cross-Office 题补齐显式 `network_mode = "no-network"`；当前 240 道题均显式关闭运行期网络。
- 为 DOCX/XLSX 的 120 个 Dockerfile 增加可覆盖的 `PIP_INDEX_URL` 与安装超时参数；这只影响镜像构建，任务运行和验证仍保持无网络。
- 为 3 道 DOCX OCR 题补齐 Tesseract；修复 4 道 DOCX Judge 对未挂载根目录 Rubric 的依赖，并校正 1 道目录页码真值。
- 为 22 个 XLSX 镜像补齐 `openpyxl`；修复 4 个 Oracle 入口路径和 6 个参考解法在 Harbor `/solution` 挂载下的运行逻辑。
- 将 29 个 XLSX Judge 从作者目录副本比较改为容器内冻结输入 SHA256 校验，并修复 2 道题正向权重总和只能达到 0.95 的问题。

## 冻结证据

逐题任务 checksum、两轮 Trial、Reward、交付文件和重复性结论见 [`docx-xlsx-final-replay-20260906.json`](docx-xlsx-final-replay-20260906.json)。本地 `jobs/` 目录不进入 Git；机器可读证据文件保留了复核所需的逐题结果。

这里冻结的是 DOCX/XLSX 候选任务包的 Oracle 可回放快照，不代表真实 Agent 实验已经完成。后续已在 [`cowork-evo-240-v0.1`](../benchmarks/cowork-evo-240-v0.1/README.md) 冻结泄漏组件和 `feedback / selection / final` 任务归属；真实 Agent Rollout 与难度校准仍需后续完成。
