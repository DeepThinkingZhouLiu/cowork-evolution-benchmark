# OfficeBench 3-19/3-20 奖学金合规案件包逐题适配记录

## 为什么合并这两道上游题

OfficeBench `3-19` 和 `3-20` 使用逐字节相同的学校政策图片和奖学金申请 PDF，分别要求转发邮件和安排复核会议。两个交付动作共享同一个合规决策，因此本题人工重构为“政策图 + 申请 PDF -> 可复算 Excel -> Word 决策 -> EML 附件草稿 + ICS 后续事项”的四产物案件包。

结构设计目标为 L3 / Hard：Agent 必须同时做多模态证据提取、政策逐条核对、证据缺口管理、多格式打包和跨文件一致性控制。正式难度保持 `unassigned`，等待冻结多 Harness Rollout 后统一经验校准。

## 实际读入上下文的材料

- `tasks/3-19/subtasks/0.json`：邮件转发要求与原 Evaluator。
- `tasks/3-20/subtasks/0.json`：合规复核会议要求与原 Evaluator。
- `fellowship_application.pdf`：申请人、邮箱、105 个英文词的 Statement of Purpose 和 1 位参考人。
- `school_policy.jpg`：逐条人工读取的 6 条申请政策。
- `utils/evaluate.py`：上游通用评测入口。
- OfficeBench 根目录 `README.md` 和 `LICENSE`，固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`。

本目录保存了两个上游 subtask、原 Evaluator 和 Apache-2.0 许可证的精确快照；Agent 输入中的 PDF 和 JPG 也是上游文件的逐字节副本。

## 逐条人工真值

| 审计对象             | 人工核对结果                                                                  |
|----------------------|-------------------------------------------------------------------------------|
| 申请人             | Alice Johnson；`alice.johnson@example.com`                                 |
| Statement of Purpose | 105 个英文词，满足不少于 100 词                                      |
| 参考人               | 只有 1 位：Dr. John Smith, Professor, University X                         |
| P01-P03              | 姓名、邮箱、最少 100 词均为 PASS                                        |
| P04                  | 政策要求 2 位参考人，实际只有 1 位，为 FAIL                            |
| P05                  | 来源没有政策截止时间或提交时间戳，为 NOT_VERIFIABLE                     |
| P06                  | 政策规定不完整申请不得考虑；本申请缺少第二位参考人，为 FAIL |
| 总体决策             | PASS 3、FAIL 2、NOT_VERIFIABLE 1；总体为 NO                         |
| 后续会议             | 上游只给定 `2024-05-31 13:00 UTC`；没有时长、地点或与会人邮箱       |

## 跨文件因果链

```text
school_policy.jpg + fellowship_application.pdf
-> 6 条政策与可追溯申请证据
-> Excel 公式计数和 NO 决策
-> 与 Excel 一致的 Word 合规报告
-> 保留原 PDF 字节的 EML 草稿
-> 只保留已知开始时间的暂定 ICS
```

Judge 不把 Golden 当作真值：它会校验冻结输入的 SHA256，重新提取 PDF 文本和政策证据，再独立检查 Excel 公式与原生图表、Word 事实、EML MIME 附件字节、ICS 时间边界、无来源断言和四产物一致性。

## 对上游的必要改造

- 保留上游全部事实和两个交付意图，但不伪造截止时间、提交时间、真实官员邮箱、会议时长、地点或与会人。
- 由于官员邮箱不可得，EML 明确使用 `.invalid` 保留地址并标记 `BLOCKED`，不把草稿伪装成可发送邮件。
- 上游 Evaluator 无法检出错误合规结论、附件篡改、无来源时间断言或跨文件矛盾。本题人工重写 18 条 Rubric 和任务专属 Judge，并加入 21 类定向负例。

## 许可、隐私与泄漏

来源为 Apache-2.0 的 OfficeBench 公开任务。按上游 benchmark 属性保守标记 `synthetic=true`。任何复用该政策图、申请 PDF 或相同决策真值的变体都必须归入 `officebench-3-19-20-fellowship-policy-application` leakage group，不得跨 split。

## 验证结果

- Golden、两次空目录 fresh replay 的 Judge Reward 均为 `1.0`；四份产物均与 Golden 逐字节一致。
- 空交付 Reward 为 `0.0`；21 个定向负例全部低于 `0.8` 通过线，最高错误得分为 `0.20`。
- XLSX 经 LibreOffice 渲染为 4 页，DOCX 渲染为 2 页，共 6 页已逐页目视检查。
- 官方 Harbor 0.21.0 Docker Oracle 独立回放两次，均为 Reward `1.0`、0 异常；两次容器产物逐字节一致。
