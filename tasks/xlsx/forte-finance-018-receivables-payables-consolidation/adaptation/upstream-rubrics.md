# Rubrics

> **通过标准：所有 rubric 均须满足，该任务才算通过。任一 rubric 不通过，则该任务整体判定为不通过。**

## 文件读取路径

```path
/workspace/solution/未收统计标准答案.csv
/workspace/input/未收统计.csv
/workspace/solution/未付统计标准答案.csv
/workspace/input/未付统计.csv
```

```json
[
  {
    "id": "01",
    "content": "输出的<file>/workspace/input/未收统计.csv</file>，与<file>/workspace/solution/未收统计标准答案.csv</file>内容一致",
    "weight": 1
  },
  {
    "id": "02",
    "content": "输出的<file>/workspace/input/未付统计.csv</file>，与<file>/workspace/solution/未付统计标准答案.csv</file>内容一致",
    "weight": 1
  },
  {
    "id": "03",
    "content": "输出的回答中，有“无僵尸账款”或其他等价表述",
    "weight": 1
  }
]
```
