# officebench-2-20-final-perfect source adaptation record

Status: `pending_human_review`; difficulty is an author hypothesis only; `calibration_status=uncalibrated`.

## Upstream evidence

- Dataset: OfficeBench, task `2-20/2`, immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original query (copied from `adaptation/source-task.json`): `Find all students from CS161_roster.csv whose final_exam is 100, and put their name and student id in makeup_exam_list.docx. Each student should be on seperate lines, with the format `student id, name`.`
- The complete upstream evaluator is retained in `adaptation/source-task.json`; its positive keywords are: `S1004`; `Dipuj`; `Fufowot`; `S1007`; `Cani`; `Jehecat`.
- Its negative keywords are: （上游 evaluator 未提供排除关键词）.
- The source license is copied to `adaptation/LICENSE-OfficeBench`.

## Output decision

The retained source evaluator names only DOCX outputs, so this package remains in `docx`. Solver deliverables are: `makeup_exam_list.docx`. No PDF, XLSX, EML, ICS, image, or TXT deliverable is silently removed. If a source task ever requires one of those formats, this script classifies it as `route_to_cross_office` instead.

## Input evidence

| Package input | SHA256 checked by this repair | Role |
|---|---|---|
| `data/input_files/CS161_Class_Roster.xlsx` | `c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7` | `solver_input` |

## Truth and Judge

Golden generation is kept separate from scoring. The source task JSON and visible package inputs are the evidence basis; the Judge reads the package inputs at evaluation time and applies the source-required terms plus task-specific structural checks. It does not compare a ZIP hash to the Golden. The original evaluator is preserved for audit only and is not treated as sufficient semantic validation.

## Known limitations and review gates

- OfficeBench fixtures are benchmark snapshots and are marked `synthetic=true`; this is a provenance label, not a claim that the source story was invented here.
- Several source families share files or near-identical queries; the `leakage_group_id` in metadata intentionally groups them for split control.
- Human review must inspect the source preview, Golden text, editable structure, pagination/rendering, and empty/corrupt/unchanged/core-error negative cases before acceptance.
