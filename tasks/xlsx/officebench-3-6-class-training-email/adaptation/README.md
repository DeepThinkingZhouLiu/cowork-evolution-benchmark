# officebench-3-6-class-training-email

状态：`pending_human_review`；L3 作者假设，`uncalibrated`。来源为 OfficeBench revision `b978b808667c32b52ce19a67ce1def1de9ae02b7` 的 `3-6/1`，输入 SHA256 `a5628f45d9f48eccd8424634fda4d3c9048200571cc03db1e6d34cf9447498f4`。原始日历/邮件副作用已离线化为一个 ICS 和四封 EML；只保留 Section A 真值并拒绝 Section B/C 邮件。

## 验证记录

已运行 `PYTHONPATH=src python -m cowork_evo_bench.cli validate-task` 和 `validate-task-assets`，均通过；Golden/fresh replay 均 `1.000000`。缺少任一交付文件、空输出或损坏输出均不通过。Judge 从固定输入排程逐文件检查，未执行无辅助 rollout，当前仍待人工复核。
