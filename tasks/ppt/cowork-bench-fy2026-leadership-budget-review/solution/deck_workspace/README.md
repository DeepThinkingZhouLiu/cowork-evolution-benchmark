# 可重建演示文稿工作区

`outline.json` 与 `truth.json` 由 `solution/solve.py` 从当前输入重新生成。稳定风格、内容、证据与资产计划分别保留在本目录。

在任务根目录执行一条命令即可重建 Golden：

```bash
PYTHON_BIN="${PYTHON_BIN:-python3.11}" \
  ./solution/solve.sh ./data/input_files ./solution ./solution/deck_workspace
```

PPTX 渲染器位于 `solution/vendor/presentation_skill/`，使用 MIT 许可；因此回放不依赖机器上的全局 Skill 路径。
