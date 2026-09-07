# officebench-3-84-sleep-analysis source adaptation record

Status: `pending_human_review`; difficulty is an author hypothesis only; `calibration_status=uncalibrated`.

## Upstream evidence

- Dataset: OfficeBench, task `3-84/0`, immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original query (copied from `adaptation/source-task.json`): `Analyze relationship between sleeping hours and age, weight, gender from survey result, convert report to report.docx file`
- The complete upstream evaluator is retained in `adaptation/source-task.json`; its positive keywords are: （上游 evaluator 未提供关键词；以原始 Query 和可见输入为真值）.
- Its negative keywords are: （上游 evaluator 未提供排除关键词）.
- The source license is copied to `adaptation/LICENSE-OfficeBench`.

## Output decision

The retained source evaluator names only DOCX outputs, so this package remains in `docx`. Solver deliverables are: `report.docx`. No PDF, XLSX, EML, ICS, image, or TXT deliverable is silently removed. If a source task ever requires one of those formats, this script classifies it as `route_to_cross_office` instead.

The upstream request asks for sleep hours and weight, but the released workbook contains neither field. Judge audit v0.2 therefore makes the smallest executable repair: it keeps the same workbook and report deliverable, changes the analysis target to the available `Satisfaction Score`, and requires age/income correlations plus gender-group means. The original request remains above and in `source-task.json` for provenance; no unavailable values are fabricated.

## Input evidence

| Package input | SHA256 checked by this repair | Role |
|---|---|---|
| `data/input_files/survey_responses.xlsx` | `fa07b47037cdd43b3218a45b9cd87d1b861a96d4097ab6facded48914ec51c11` | `solver_input` |

## Truth and Judge

Golden generation is kept separate from scoring. The source task JSON and visible package inputs are the evidence basis; the Judge reads the package inputs at evaluation time and applies the source-required terms plus task-specific structural checks. It does not compare a ZIP hash to the Golden. The original evaluator is preserved for audit only and is not treated as sufficient semantic validation.

## Known limitations and review gates

- OfficeBench fixtures are benchmark snapshots and are marked `synthetic=true`; this is a provenance label, not a claim that the source story was invented here.
- Several source families share files or near-identical queries; the `leakage_group_id` in metadata intentionally groups them for split control.
- Human review must inspect the source preview, Golden text, editable structure, pagination/rendering, and empty/corrupt/unchanged/core-error negative cases before acceptance.
- This adaptation is intentionally not comparable to a hypothetical sleep/weight analysis; it evaluates whether the agent notices the released schema and performs the explicitly repaired analysis.
