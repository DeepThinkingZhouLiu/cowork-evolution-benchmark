# automation-vendor-recon（R044 / XLSX-C090）

状态：`pending_human_review`；作者难度假设 `L3`，`calibration_status=uncalibrated`。

## 来源审计

- 上游：AutomationBench `automationbench/domains/finance/tasks.py`，固定 revision `4a8e1061254004d9dac807054eed33fad7d1ff14`，任务函数 `get_fin_vendor_statement_recon_task`，`example_id=4075`。
- 许可：上游仓库 MIT；许可文本保留在 `LICENSE-AutomationBench`。该任务只使用代码内嵌公开 mock 状态，不含真实个人或企业数据。
- 上游状态与原始请求逐字保存于 `upstream-task.json`。原始 assertions 要求 controller 汇总、Acme 2,000 差异、Metro 780 差异，且不得联系已匹配 CloudHost 或明确排除的 TechServe。

## 离线改造与证据账本

| 要求/事实 | 证据 | 改造/验证 |
|---|---|---|
| 四家供应商及邮箱、声明余额 | 上游 `Vendor Statements Received` | 输入工作表同名保留；Judge 读取 |
| AP 余额和 payment-in-transit 备注 | 上游 `Our AP Balances` | 输入工作表同名保留；Judge 读取 |
| 差额=声明余额−AP 余额 | 上游用户 prompt 与 assertions | `Reconciliation` 逐行重算 |
| Acme 2,000、Metro 780 为未解决差异 | 上游 assertions；由可见余额独立算术验证 | Judge 仅从输入计算 |
| CloudHost 匹配、TechServe 排除 | 上游正/负 assertions及备注 | Judge 检查状态、差额、联系标记和邮件收件人 |
| Gmail 汇总和供应商联系 | 上游 assertions | `Email Draft` 表示本地可审计邮件，不执行外部副作用 |

原始 Google Sheets/Gmail workflow 被压缩成一个可编辑 XLSX：保留两个来源表，新增 `Reconciliation` 和 `Email Draft`。未引入来源之外的事实；金额以数值存储、可编辑并显示千位分隔符。

## 验证与交接

`solution/solve.py` 从输入重新读取并生成 Golden；`tests/judge.py` 独立读取输入并重算，不比较 Golden 哈希。已设计文件缺失、损坏、只复制输入、漏掉差异、错误排除和错误联系等负例。正式运行记录由制题交接时补充；当前不得标记 `accepted` 或 `released`。
