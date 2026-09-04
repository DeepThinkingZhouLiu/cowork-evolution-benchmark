# officebench-1-15-project-report-reference source adaptation record

Status: `pending_human_review`; difficulty is an author hypothesis only; `calibration_status=uncalibrated`.

## Upstream evidence

- Dataset: OfficeBench, task `1-15/2`, immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original query (copied from `adaptation/source-task.json`): `Retrieve reference text from reference.txt and add it to the end of word file project_report.docx.`
- The complete upstream evaluator is retained in `adaptation/source-task.json`; its positive keywords are: `Martin Arjovsky`; `Amar Shah`; `Yoshua Bengio`; `Unitary Evolution Recurrent Neural Networks`.
- Its negative keywords are: （上游 evaluator 未提供排除关键词）.
- The source license is copied to `adaptation/LICENSE-OfficeBench`.

## Output decision

The retained source evaluator names only DOCX outputs, so this package remains in `docx`. Solver deliverables are: `project_report_edited.docx`. No PDF, XLSX, EML, ICS, image, or TXT deliverable is silently removed. If a source task ever requires one of those formats, this script classifies it as `route_to_cross_office` instead.

## Input evidence

| Package input | SHA256 checked by this repair | Role |
|---|---|---|
| `data/input_files/project_report.docx` | `df8232537be6927a7cf34a0caf8cb20b1f9be3c7d73596648b90704caad176b0` | `solver_input` |
| `data/input_files/reference.txt` | `032c567544cc48bb9b1f03e19444be48b8f275b61bf634cdd9f8d757bfde8bbe` | `solver_input` |

## Truth and Judge

Golden generation is kept separate from scoring. The source task JSON and visible package inputs are the evidence basis; the Judge reads the package inputs at evaluation time and applies the source-required terms plus task-specific structural checks. It does not compare a ZIP hash to the Golden. The original evaluator is preserved for audit only and is not treated as sufficient semantic validation.


## Author-side validation (2026-09-04)

- `validate-task` and `validate-task-assets`: pass.
- Checked-in Golden and a fresh replay from packaged inputs: reward `1.000000`; each was scored five times with identical deterministic results.
- Missing output and corrupt bytes: fail, with reward between `0.000000` and `0.100000`. A screenshot-only DOCX also fails.
- A non-empty targeted defect that removes one required fact, paragraph, or table row: reward `0.15` and fail. For source-edit tasks, the unchanged input also fails.
- LibreOffice `7.3.7.2` rendered Golden and fresh replay to 5 pages each. Text and pixel-boundary checks found no unintended blank page or clipping, and the rendered contact sheet was inspected.

These are author-side checks, not human acceptance or rollout calibration. Status remains `pending_human_review`; `calibration_status=uncalibrated`.
## Known limitations and review gates

- OfficeBench fixtures are benchmark snapshots and are marked `synthetic=true`; this is a provenance label, not a claim that the source story was invented here.
- Several source families share files or near-identical queries; the `leakage_group_id` in metadata intentionally groups them for split control.
- Human review must inspect the source preview, Golden text, editable structure, pagination/rendering, and empty/corrupt/unchanged/core-error negative cases before acceptance.
