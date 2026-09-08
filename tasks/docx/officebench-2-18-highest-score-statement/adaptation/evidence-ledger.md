# 证据台账

| 来源 | SHA256 | 贡献 | 阅读结论 |
|------|--------|------|----------|
| `nomination_shortlist.xlsx` | `4a70c6310035b9b40d7a5012bdc91cc72bd539a5cd313fe7672e09e1ec65b598` | 4 行提名短表 | 最高列值 Rachel=100，不代表完整班级范围 |
| `detailed_midterm_register.xlsx` | `a5c8026dc09f38411c6b8783dad74f7f66f08387eeb21b9d8eb7f28a90dd5b6e` | 20 行详细记录 | Tina 与 Rachel 在 Midterm 1 并列 100；中间存在 4 组重名 |
| `legacy-officebench-source-task.json` | `29189cef7323f16d7ee5ff6c6c7e5c09e8fea4c7e240ee07d6664ed3a54814a3` | 原 Word 陈述 Query/Evaluator | 仅检查 `Rachel.docx` 存在，已被 v0.3 任务取代 |
| `source-score-audit-spread.json` | `a7ab25fadaa5cab5b36d2f0eeac84903ad41806ac62cb67078f67a537678c36f` | 上游极差 Query/Evaluator | 关键值 40，新 Judge 从 XLSX 重算 |
| `source-score-audit-declines.json` | `c644daae5d4ccc87b09bec6ceb6ae3ca08015c7e30ae7f4173e5843c974c431b` | 上游跨期比较 Query/Evaluator | 确认两次成绩的比较用途 |
| `LICENSE-OfficeBench` | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4` | 派生/发布许可 | Apache-2.0 |

遗留字节已以 `legacy-officebench-{data,environment}-*.xlsx`、`legacy-officebench-{golden,solution}.docx` 冻结保存，用于追溯而不会复制到 `/app/data/input_files`。原 data/environment 成对文件的 SHA 一致；原 Golden/solution 均为 `750c92fd01340e548f48783b08b7a93b4f45596f615602478eef3cc97af97d77`。
