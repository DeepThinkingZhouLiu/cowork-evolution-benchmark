# 验证记录

验证日期为 2026-08-24。以下命令均从仓库根目录执行；`$TASK` 指向本题目录，`$PY` 指向项目专用 Conda 环境的 Python。

## 来源与输入

```bash
UPSTREAM=REMOVED_LOCAL_PATH
$PY "$TASK/adaptation/extract_snapshot.py" \
  --dump "$UPSTREAM/db/init.sql.gz" \
  --output /tmp/EMPLOYEES.copy.sql
cmp /tmp/EMPLOYEES.copy.sql \
  "$TASK/data/input_files/HR_ANALYTICS__PUBLIC__EMPLOYEES.copy.sql"
```

结果：固定上游 dump 哈希通过；原样抽取 50000 行、18 列；输出 SHA256 为 `c1dfba6c1b0c55b03e47cb46d093202af2befb9593bc7f4cf4796a9781844993`；`cmp` 通过。adaptation 中保存的 Query、system prompt、task config、preprocess、Evaluator、PDF、两份上游 Golden 和 LICENSE 也分别与上游 commit 执行了 `cmp`，全部通过。

## 结构、Golden 与回放

```bash
PYTHONPATH=src $PY -m cowork_evo_bench.cli validate-task "$TASK"
PYTHONPATH=src $PY -m cowork_evo_bench.cli validate-task-assets "$TASK"
$PY "$TASK/tests/judge.py" \
  --artifact-dir "$TASK/solution" \
  --data-dir "$TASK/data/input_files"
$PY "$TASK/tests/replay_tests.py" \
  --task-root "$TASK" \
  --result "$TASK/metadata/replay-report.json"
```

结果：两个仓库校验器均通过；Golden 为 1.0；fresh replay 为 1.0；内容变异 replay 为 1.0；8 个针对性负例全部被拒绝。完整负例结果见 `replay-report.json`。

## Office 渲染与目视检查

```bash
$PY REMOVED_LOCAL_PATH \
  --input "$TASK/solution/budget_vs_actuals.pptx" --json
libreoffice --headless --convert-to pdf --outdir /tmp/ppt-render \
  "$TASK/solution/budget_vs_actuals.pptx"
libreoffice --headless --convert-to pdf --outdir /tmp/xlsx-render \
  "$TASK/solution/budget_vs_actuals.xlsx"
```

结果：presentation-skill 的视觉门禁返回空问题列表 `[]`。PPTX 渲染为 10 页 16:9 PDF；XLSX 渲染为 2 页 A4 PDF。12 页都已逐页打开目视检查，记录和最终文件哈希见 `visual-inspection.md`。

## Docker 离线回放

Dockerfile 默认使用 Debian 官方镜像。由于本机官方镜像经代理下载很慢，构建验证使用 Dockerfile 暴露的可选参数切换为清华镜像；依赖版本与默认构建完全相同。

```bash
docker build --network=host \
  --build-arg DEBIAN_MIRROR=https://mirrors.tuna.tsinghua.edu.cn/debian \
  --build-arg DEBIAN_SECURITY_MIRROR=https://mirrors.tuna.tsinghua.edu.cn/debian-security \
  --build-arg HTTP_PROXY=http://127.0.0.1:7890 \
  --build-arg HTTPS_PROXY=http://127.0.0.1:7890 \
  -t cowork-evo-budget-review:test \
  -f "$TASK/environment/Dockerfile" "$TASK/environment"

mkdir -p /tmp/cowork-budget-docker-output
docker run --rm --network=none \
  -v "$TASK:/task:ro" \
  -v /tmp/cowork-budget-docker-output:/app/output \
  cowork-evo-budget-review:test \
  sh -lc 'python /task/solution/solve.py --input-dir /task/data/input_files --output-dir /app/output --workspace-dir /tmp/workspace && python /task/tests/judge.py --artifact-dir /app/output --data-dir /task/data/input_files'
```

结果：镜像 `sha256:27a7cf52715e016fb1e1dec7ed0ad99432b853685ecabfcbabd1a0fea2e7b23a` 构建成功；在 `--network=none` 下重新生成两个交付物并取得 1.0。镜像内为 Python 3.11.16、LibreOffice 7.4.7.2、pdftotext 22.12.0、lxml 6.0.1、openpyxl 3.1.5、Pillow 11.3.0、python-pptx 1.0.2。
