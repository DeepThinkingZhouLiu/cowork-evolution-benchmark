# 公开任务生成工具

本目录只负责机械化生成无隐私的公开端到端样例，包括输入文件、参考解法、Golden、任务级 Judge 和来源哈希。它不负责把样例自动晋升为正式 Benchmark 题，也不替代人工视觉复核、Judge 校准和多模型难度 Rollout。

## 环境

本项目使用独享 Conda 环境：

```text
/mnt/bn/liuzhou-hl-training/liuzhou/conda_envs/cowork-evolution-benchmark
```

可复现依赖记录在 `environment.yml` 和 `requirements.txt`。生成器只使用项目自己合成的业务材料，不执行 `003-Reference` 中任何外部仓库脚本。

## 文件职责

| 文件                                  | 责任                                                         |
|---------------------------------------|--------------------------------------------------------------|
| `common.py`                           | Office 文件构造、Harbor 包骨架、通用确定性 Judge 和来源哈希   |
| `generate_docx_tasks.py`              | 生成 8 个 Word 任务家族                                      |
| `generate_cross_office_tasks.py`      | 生成 8 个跨 Office 任务家族                                  |
| `generate_ppt_tasks.py`               | 生成 12 个 PPT 任务家族                                      |
| `generate_xlsx_tasks.py`              | 生成 14 个 Excel 任务家族                                    |
| `build_catalog.py`                    | 扫描全部任务并生成统一机器目录                               |

## 发布边界

- 生成器输出是可运行的公开样例，不是论文最终数据规模。
- 合成数据必须标记 `synthetic = true`，不能冒充真实用户轨迹。
- 参考仓库只提供任务形态、Rubric 和 Judge 设计启发，衍生任务不得复制许可不明或带隐私的内容。
- 正式 Release 仍需执行隐私与许可审查、近重复聚类、Golden 人审、正负样本校准、Judge 稳定性和难度筛选。
