# sb2-template-14-09-concession-revenue

状态：`pending_human_review`；难度为 `L2` 作者假设，`calibration_status=uncalibrated`。

## 来源与证据

- 来源：SpreadsheetBench-2 sample，固定 revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece`。
- 输入：`data/input_files/concession_input.xlsx`，SHA256 `b272c7de3b687ecea8013d95e57cdeeccb071c3e7382d1a70abdff7b4216ec69`，角色为 solver-visible input。
- 上游 Golden 原件保留为 `adaptation/upstream-golden-concession_revenue.xlsx`（仅作证据）。它将 Years 3–5 concessions 写成硬编码数值 `0`，且没有设置自动重算标志；这与当前要求的原生可编辑公式合同不一致。依据输入和规则重新生成 reviewed Golden `solution/Golden/concession_revenue.xlsx`，其中三格使用 `=0` 并启用自动重算。本地 Judge 不读取 Golden，而是从 solver-visible 模板和明确规则独立生成 72 个目标公式。
- 许可：CC-BY-SA-4.0；当前按公开候选处理，最终发布仍需人工复核。

## 改造边界

保留原始工作流和输入，离线化运行环境，明确输出路径 `/app/output/concession_revenue.xlsx`，要求原生可编辑工作簿、公式/值保真和非目标结构保持。市场租金、Loss-to-Lease、其他收入、汇总和 vacancy 公式均从模板假设单元格推导；Years 3–5 concessions 按指令规则写为 `=0`。未加入 Golden 之外的业务事实。

## Evidence-to-requirement map

| Requirement | Evidence | Judge |
|---|---|---|
| Five-year market rent and other-income projections | visible assumption cells C7, C11, C17:C24 and year headers | R002/R003 |
| Loss-to-Lease and general vacancy calculations | visible rates C12 and C27 plus row labels | R002/R003 |
| Years 3–5 concessions equal zero | adapted instruction and visible Year 1/2 concession assumptions | R002 |
| Native formulas and template preservation | input workbook structure and editable-cell requirement | R004/R005 |
| Input provenance | input-manifest SHA256 | R006 |

## 验证记录

已运行 `PYTHONPATH=src python -m cowork_evo_bench.cli validate-task` 和 `validate-task-assets`；本题均通过。reviewed Golden 与 fresh replay 均经独立 Judge 得到 `1.000000`；空/损坏输出均 `-0.150000`，未修改输入和单个公式错误均 `0.150000`，模板外 collateral 改动为 `0.800000` 但因 R004 hurdle 不通过。LibreOffice 重算目标数值无差异，并成功渲染 `concession_revenue.pdf`。Judge 覆盖文件完整性、输入推导公式、LibreOffice 可重算数值、非目标保持、结构解析和输入哈希。尚未执行无辅助模型 rollout；难度和发布状态仍需人工复核。

## 已知限制

上游工作簿可能含 Excel 专有公式或图表行为；离线环境只验证 OOXML 结构和确定性目标，不替代人工打开、重算和视觉检查。当前候选不标记为 accepted 或 released。
