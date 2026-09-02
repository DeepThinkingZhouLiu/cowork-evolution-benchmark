# AutomationBench 5011 逐题改造记录

- 上游：`get_hr_salary_band_audit_task()` / `hr.salary_band_audit`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `b0d32626b11150c8adb95f24a4b550ae540849b9d49df95c9974c35f8ba30a20`。
- 已完整读取原 Query、5 档带宽、6 人薪酬、2 条 Flag 惯例和 8 条断言；`materialize_inputs.py` 只处理 5011。
- 人工真值：Bob 125000 相对 IC2 上界 120000 为 Over 5000；Carol 108000 相对 IC3 下界 110000 为 Under -2000；Eve 215000 相对 IC4 上界 210000 为 Over 5000；其余三人 Within。
- Golden 含 10 张 Excel 表、72 个公式、3 页 Word、3 页 PPT、1 封 MBOX 草稿和 2 份控制 JSON。Judge 独立核验带宽查找、异常、金额不舍入、零更新与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；16 页目视检查通过。
