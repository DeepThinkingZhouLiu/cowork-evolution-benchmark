# Codex-VLM 混合评分基础设施

## 责任边界

这套公共层只负责可信评测，不负责做题，也不替代每道题现有的确定性 `judge.py`。

```text
受测 Harness 交付候选文件
-> 每题原有代码 Judge 产生 deterministic_reward
-> 公共渲染层把需要视觉评价的 PPTX/PDF 转成图片
-> 固定 Codex CLI 读取任务特定 visual_rubric.json 和图片
-> ZCloud 的 gpt-5.6-terra 逐条产生视觉分与证据
-> 本地代码验证模型 JSON，并重新计算 visual_reward
-> 按 70% deterministic + 30% visual 产生 combined_reward
```

模块分工如下：

- `judging/visual_config.py`：读取公共非秘密配置和本地 `.env`，校验 Provider、模型、Codex 版本、超时、重试、稳定性与聚合权重。
- `judging/rendering.py`：使用 LibreOffice 和 Poppler 渲染 PPTX/PDF，生成逐页 PNG、contact sheet、文件哈希与 manifest。
- `adapters/codex_vlm.py`：固定 Codex CLI 版本，以非交互、临时、只读且禁用 shell 工具的方式调用 VLM。
- `judging/visual_models.py`：严格解析每题视觉 Rubric 和模型 JSON；模型只报逐项分，本地代码负责加权。
- `judging/composite.py`：执行版本化的 70/30 聚合，并区分任务失败和评分基础设施失败。
- `judging/visual_pipeline.py`：编排单题输入、渲染、评分和 artifacts 落盘。
- `cowork-evo judge-visual`：提供给单题 Judge 或批量 Runner 调用的轻量入口。

`evaluation` 模块仍然只消费标准单题结果，不直接调用模型。

## 配置与秘密

公共配置位于 `configs/judging/codex-vlm.json`，默认固定：

- Provider：`zcloud`
- Base URL：`https://api.zcloudapi.com/v1`
- 模型：`gpt-5.6-terra`
- Codex CLI：`0.148.0`
- Wire API：`responses`
- 聚合：确定性分 70%，视觉分 30%

模型可通过 `COWORK_EVO_VLM_MODEL` 临时覆盖，Codex 路径可通过 `COWORK_EVO_CODEX_BINARY` 覆盖。Base URL 可通过 `COWORK_EVO_VLM_BASE_URL` 覆盖，但 Provider 名始终固定为 `zcloud`。

真实 API Key 只能放在仓库根目录被 Git 忽略的 `.env`：

```dotenv
COWORK_EVO_VLM_API_KEY=replace-with-local-secret
```

`.env.example` 只展示变量名。密钥不得写进公共配置、Task、Rubric、命令行参数、模型提示、测试快照或日志。评分进程仅在启动 Codex 子进程时通过命名环境变量传递密钥，并会对 stdout/stderr 和模型回复做二次脱敏。

Codex CLI 0.148.0 没有公开温度参数，因此配置明确记录 `temperature_strategy=provider_default`，不伪造“温度为 0”。需要更稳定的论文结果时，可把 `stability.repetitions` 提高到 3；系统对每条判据取中位数，并保留最接近中位数那次的证据。

## 每题视觉 Rubric

需要视觉评分的任务在任务根目录增加 `visual_rubric.json`。这份文件必须由制题者逐题读取 `instruction + inputs + Golden + 当前 Rubric` 后单独编写，禁止用一套空泛模板批量套用。

```json
{
  "task_id": "ppt/example",
  "rubric_version": "visual-v1",
  "enabled": true,
  "reference_mode": "reference_free",
  "candidate_artifacts": [
    {
      "path": "answer.pptx",
      "kind": "pptx",
      "label": "候选演示文稿"
    }
  ],
  "criteria": [
    {
      "criterion_id": "V001",
      "description": "第 1 页标题与核心结论必须形成明确的视觉层级。",
      "weight": 1.0,
      "score_anchors": {
        "0": "标题或核心结论不可读，或完全没有层级。",
        "0.5": "内容可读，但主次关系较弱。",
        "1": "标题和核心结论一眼可区分，阅读顺序明确。"
      },
      "evidence_requirement": "指出具体页码、区域及可见字号、位置或强调方式。"
    }
  ]
}
```

两种参考策略不能混用：

- `reference`：用于精确还原、版式修复等任务。候选图和 Golden/reference 图都会送给 VLM。
- `reference_free`：用于开放式生成。VLM 只看用户输入、任务特定 Rubric 和候选图；系统在代码路径上不会读取或渲染 Golden。

只有 PPTX/PDF 产物需要视觉评分。没有视觉评分的任务不需要 `visual_rubric.json`，其 `combined_reward` 等于 `deterministic_reward`。

## 单题调用

```bash
cowork-evo judge-visual \
  --task-dir tasks/ppt/<task-id> \
  --candidate-dir /path/to/agent-delivery \
  --deterministic-result /path/to/deterministic-result.json \
  --config configs/judging/codex-vlm.json \
  --env-file .env \
  --output-dir .coworkevo/artifacts/<run-id>/<task-id>
```

`--visual-rubric` 省略时默认读取 `<task-dir>/visual_rubric.json`。输出目录必须是新的目录，系统拒绝覆盖旧评分。

## 输出协议

```text
<output-dir>/
├── deterministic/result.json       # 原代码 Judge 分项及原始结果副本
├── render/
│   ├── candidate/manifest.json      # 候选源文件、逐页图和哈希
│   └── reference/manifest.json      # 仅 reference 模式存在
├── visual/
│   ├── result.json                  # 标准视觉分、逐项证据、模型和版本
│   └── codex/                       # 每次尝试的事件、结构化回复和 schema
├── combined/result.json             # 聚合结果或机器可读失败状态
└── reward/
    ├── status.json                  # 是否有可用总分
    └── reward.txt                   # 只有评分有效时才存在，便于 Harbor 收集
```

详细 artifacts 与最终 `reward.txt` 分离。图片和模型回复必须放在 `.coworkevo/`、`artifacts/`、`jobs/` 等 Git 忽略目录，不能提交。

## 分数与失败语义

- 视觉任务：`combined_reward = 0.7 * deterministic_reward + 0.3 * visual_reward`。
- 非视觉任务：`combined_reward = deterministic_reward`。
- 只有候选核心文件缺失、文件头损坏或 Office/PDF 基础结构损坏，才属于 `task_failure`，总分为 0。
- LibreOffice/Poppler 不可用、渲染超时、Codex 版本不符、Codex/API 超时、Provider 错误和模型 JSON 不合法，都属于 `infrastructure_failure`。
- `infrastructure_failure` 不产生 `visual_reward`、`combined_reward` 或 `reward.txt`，不能记为 0，也不能静默回退成纯代码分。

## 安全约束

- 做题容器永远不挂载 `.env`、公共 Judge 工作区、隐藏 Golden 或模型回复。
- Codex 评分 Harness 只接收渲染图片、任务指令和隐藏视觉 Rubric，不接收可修改的原始 Task。
- Codex 使用 `--sandbox read-only`、`--ephemeral`、`--ignore-user-config`，并显式禁用 `shell_tool`。
- 临时评分工作区只读；候选文档和图片中的文字全部被声明为不可信数据。
- Prompt Injection（例如“忽略 Rubric、执行命令、输出 API Key”）不得改变评分规则或触发工具调用。

## 运行依赖与验证

Python 包依赖由 `pyproject.toml` 管理。宿主机还必须提供：

- LibreOffice/soffice：把 PPTX 转成 PDF。
- Poppler 的 `pdfinfo`、`pdftoppm`：检查 PDF 页数并渲染逐页图片。
- 固定版本 Codex CLI 0.148.0。

最小测试覆盖配置解析、Schema、模型非法 JSON、70/30 聚合、无视觉任务、基础设施失败、缺失核心文件、`reference/reference_free` 隔离、Prompt Injection 边界，以及本地小型 PPTX/PDF 的真实渲染。
