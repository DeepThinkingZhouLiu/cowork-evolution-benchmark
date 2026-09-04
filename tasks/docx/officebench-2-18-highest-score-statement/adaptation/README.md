# officebench-2-18-highest-score-statement source adaptation record

Status: `pending_human_review`; difficulty is an author hypothesis only; `calibration_status=uncalibrated`.

## Upstream evidence

- Dataset: OfficeBench, task `2-18/0`, immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original query (copied from `adaptation/source-task.json`): `Write a statement in docx file to the student with the highest score in the class, use the student's name as filename`
- The complete upstream evaluator is retained in `adaptation/source-task.json`; its positive keywords are: （上游 evaluator 未提供关键词；以原始 Query 和可见输入为真值）.
- Its negative keywords are: （上游 evaluator 未提供排除关键词）.
- The source license is copied to `adaptation/LICENSE-OfficeBench`.

## Output decision

The retained source evaluator names only DOCX outputs, so this package remains in `docx`. Solver deliverables are: `highest_score_student.docx`. No PDF, XLSX, EML, ICS, image, or TXT deliverable is silently removed. If a source task ever requires one of those formats, this script classifies it as `route_to_cross_office` instead.

## Input evidence

| Package input | SHA256 checked by this repair | Role |
|---|---|---|
| `data/input_files/class_member.xlsx` | `0cba4e5795d7eaec5f6a3cb8d4725f12ca4e9935dea04a32e4ed21ce162a7493` | `solver_input` |
| `data/input_files/score.xlsx` | `4a70c6310035b9b40d7a5012bdc91cc72bd539a5cd313fe7672e09e1ec65b598` | `solver_input` |
| `data/input_files/shopping_list.xlsx` | `e6468e5452830bc4bd8767fb9a95c0ea95a65d876e306581437889e43c2a821a` | `solver_input` |

## Truth and Judge

Golden generation is kept separate from scoring. The source task JSON and visible package inputs are the evidence basis; the Judge reads the package inputs at evaluation time and applies the source-required terms plus task-specific structural checks. It does not compare a ZIP hash to the Golden. The original evaluator is preserved for audit only and is not treated as sufficient semantic validation.

## Known limitations and review gates

- OfficeBench fixtures are benchmark snapshots and are marked `synthetic=true`; this is a provenance label, not a claim that the source story was invented here.
- Several source families share files or near-identical queries; the `leakage_group_id` in metadata intentionally groups them for split control.
- Human review must inspect the source preview, Golden text, editable structure, pagination/rendering, and empty/corrupt/unchanged/core-error negative cases before acceptance.
