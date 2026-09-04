# 逐题改造说明

本题不是本项目自编的业务题。原 Query、3 份往来明细 XLSX、2 份标准答案 CSV 和原始 rubric 全部来自 FORTE 公开仓库的 `Finance-018`，并固定到 commit `345c1ec1487139db9dd319787fa9405ba85d1869`。

## 上游原件

| 本地审计文件                              | 上游内容                           | SHA256                                                           |
|-------------------------------------------------|------------------------------------|------------------------------------------------------------------|
| `upstream-task.md`                              | 原 Query、输入清单、rubric         | `01506a107ae61814e220cd6aca06605172be3e8687005fc1e275813120baf393` |
| `upstream-rubrics.md`                           | 原始评分标准                       | `5f1b1f7af5e06d9de60dfc420e3235eda1ebe1d94f8163682752321e2385f61c` |
| `upstream-unpaid-golden.csv`                    | 上游未付标准答案                   | `ecc14fd234ebecb51380e3deab5e459871f777d0f9de858df546f57c15eb10d5` |
| `upstream-unreceived-golden.csv`                | 上游未收标准答案                   | `8d4a9c0e988afa7099512e9ed0c6eb041bc0a11da7baa286041e2d2c81c8d6d3` |
| `upstream-finance-data-analyzer-SKILL.md`       | 上游随题财务分析 Skill              | `8d135efe6836cb42bb47937e5903a42940e914222f8b7edbe88b4a63e61ceea4` |
| `upstream-accounting-rules.md`                  | 上游会计规则参考                     | `61d16a08d55ef13452bc61dfabb24112c2af6136c2af288c07e399be8cb06550` |
| `upstream-data-cleaning-rules.md`               | 上游数据清洗规则参考                 | `af5f61dc7f4ef3641f6336caa2214f620ce82a3a50e18f0bed111e8ae540f411` |
| `LICENSE-FORTE`                                 | MIT 许可证                         | `0746b4b0c31d032479c7fdf1aa17566eb1978392126bcde26eefcfceb6893f9c` |

三份 Agent 可见输入也是从该 commit 原样复制，具体哈希和字节数见 `metadata/input-manifest.json`。

## 人工逐文件审计

三份工作簿都只有一张表，且共用同一个 10 列表头。去除空行后，2025 上半年、2025 下半年、2026 年分别有 74、59、58 条明细，合计 191 条。

根据原 Query 的“目前”口径，人工核对 2026 年期末方向与余额：

- 期末方向为“贷”且余额非零的未付共 31 笔，合计 `3,984,606.46`。
- 期末方向为“借”且余额非零的未收共 2 笔，合计 `4,992,891.47`。
- 绵阳长城发展融资担保有限公司的三期余额为 `1,500,000.00 -> 1,500,000.00 -> 1,700,000.00`，已变动。
- 黄杉文化传播有限公司的三期余额为 `1,853,858.43 -> 4,672,887.50 -> 3,292,891.47`，已变动。
- 因此原始数据中没有连续三期余额不变的僵尸账款，与上游第 03 条 rubric 一致。

本地处理器从工作簿现算这些结果，生成的两份 CSV 与上游 GB18030/CRLF 标准答案逐字节相同。

## 改造边界

- 原 Query 要求把两份 CSV 写到原 workspace，并在聊天中说明僵尸账。Harbor v0.2.0 只要求 `/app/output/往来款跨期整合.xlsx`：两份当期统计、三期僵尸账核查和来源校验摘要均在该工作簿的可编辑工作表中呈现。
- `reconcile_workbooks.py`、上游 CSV 和文本回答仍保留在参考/评测侧，作为独立复算与变异测试证据，不属于 solver 交付物。
- 处理器不依赖上游 Golden。Judge 会修改一笔应收余额使其成为僵尸账，并新增一笔应付；只复制固定成品或写死原金额会被拒绝。

## 发布声明

FORTE 公开 README 把该项目定位为办公真实任务评测，但并未公开证明 Finance-018 的往来明细来自某家真实企业生产系统。因此本项目保守地标记 `synthetic=true`：它是真实存在的上游 benchmark 题，但不将其包装成真实公司财务记录。
