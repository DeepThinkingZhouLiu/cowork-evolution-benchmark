# officebench-2-26-market-analysis-text source adaptation record

Status: `pending_human_review`; difficulty is an author hypothesis only; `calibration_status=uncalibrated`.

## Upstream evidence

- Dataset: OfficeBench, task `2-26/1`, immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original query (copied from `adaptation/source-task.json`): `Retrieve the text from market_analysis.pdf and insert it into a newly created market_analysis.docx file`
- The complete upstream evaluator is retained in `adaptation/source-task.json`; its positive keywords are: `Market`; `Analysis`; `Marketing`; `Plan`; `Business`; `Customer`; `How Large`.
- Its negative keywords are: （上游 evaluator 未提供排除关键词）.
- The source license is copied to `adaptation/LICENSE-OfficeBench`.

## Output decision

The retained source evaluator names only DOCX outputs, so this package remains in `docx`. Solver deliverables are: `market_analysis.docx`. No PDF, XLSX, EML, ICS, image, or TXT deliverable is silently removed. If a source task ever requires one of those formats, this script classifies it as `route_to_cross_office` instead.

## Input evidence

| Package input | SHA256 checked by this repair | Role |
|---|---|---|
| `data/input_files/market_analysis.pdf` | `8e39f62f9fb5747c5ac1527eea0162ef7f91a5f178a9890eb46a995ef5d5207c` | `solver_input` |

## Truth and Judge

Golden generation is kept separate from scoring. The source task JSON and visible package inputs are the evidence basis; the Judge reads the package inputs at evaluation time and applies the source-required terms plus task-specific structural checks. It does not compare a ZIP hash to the Golden. The original evaluator is preserved for audit only and is not treated as sufficient semantic validation.

## Known limitations and review gates

- OfficeBench fixtures are benchmark snapshots and are marked `synthetic=true`; this is a provenance label, not a claim that the source story was invented here.
- Several source families share files or near-identical queries; the `leakage_group_id` in metadata intentionally groups them for split control.
- Human review must inspect the source preview, Golden text, editable structure, pagination/rendering, and empty/corrupt/unchanged/core-error negative cases before acceptance.
