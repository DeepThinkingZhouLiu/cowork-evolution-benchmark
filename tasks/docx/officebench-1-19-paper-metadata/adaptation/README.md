# officebench-1-19-paper-metadata source adaptation record

Status: `pending_human_review`; difficulty is an author hypothesis only; `calibration_status=uncalibrated`.

## Upstream evidence

- Dataset: OfficeBench, task `1-19/0`, immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original query (copied from `adaptation/source-task.json`): `extract title and authors of the paper from pdf file and save in word file named 'paper.docx'`
- The complete upstream evaluator is retained in `adaptation/source-task.json`; its positive keywords are: `OS-COPILOT`; `Zhiyong`; `Wu`; `Chengcheng`; `Han`; `Zichen`; `Ding`.
- Its negative keywords are: （上游 evaluator 未提供排除关键词）.
- The source license is copied to `adaptation/LICENSE-OfficeBench`.

## Output decision

The retained source evaluator names only DOCX outputs, so this package remains in `docx`. Solver deliverables are: `paper.docx`. No PDF, XLSX, EML, ICS, image, or TXT deliverable is silently removed. If a source task ever requires one of those formats, this script classifies it as `route_to_cross_office` instead.

## Input evidence

| Package input | SHA256 checked by this repair | Role |
|---|---|---|
| `data/input_files/paper.pdf` | `ab6d2a7e041b267c07e064c4acd1c65c557969b1a440e80e5b2ecbab70cbf228` | `solver_input` |

## Truth and Judge

Golden generation is kept separate from scoring. The source task JSON and visible package inputs are the evidence basis; the Judge reads the package inputs at evaluation time and applies the source-required terms plus task-specific structural checks. It does not compare a ZIP hash to the Golden. The original evaluator is preserved for audit only and is not treated as sufficient semantic validation.


## Author-side validation (2026-09-04)

- `validate-task` and `validate-task-assets`: pass.
- Checked-in Golden and a fresh replay from packaged inputs: reward `1.000000`; each was scored five times with identical deterministic results.
- Missing output and corrupt bytes: fail, with reward between `0.000000` and `0.100000`. A screenshot-only DOCX also fails.
- A non-empty targeted defect that removes one required fact, paragraph, or table row: reward `0.69` and fail. For source-edit tasks, the unchanged input also fails.
- LibreOffice `7.3.7.2` rendered Golden and fresh replay to 1 page each. Text and pixel-boundary checks found no unintended blank page or clipping, and the rendered contact sheet was inspected.

These are author-side checks, not human acceptance or rollout calibration. Status remains `pending_human_review`; `calibration_status=uncalibrated`.
## Known limitations and review gates

- OfficeBench fixtures are benchmark snapshots and are marked `synthetic=true`; this is a provenance label, not a claim that the source story was invented here.
- Several source families share files or near-identical queries; the `leakage_group_id` in metadata intentionally groups them for split control.
- Human review must inspect the source preview, Golden text, editable structure, pagination/rendering, and empty/corrupt/unchanged/core-error negative cases before acceptance.
