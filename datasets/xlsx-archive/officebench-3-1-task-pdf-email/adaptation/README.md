# C018 adaptation record

Status: `pending_human_review`. The `advanced` tier is an L3 author hypothesis, not a calibrated release label.

## Source and evidence

- Source: OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`, `tasks/3-1/subtasks/0.json`.
- Original request: `Convert Excel spreadsheet containing project tasks to task.pdf and email Alice.eml containing pdf file for review`.
- Input: `tasks/3-1/testbed/data/task.xlsx`, 11,577 bytes, SHA256 `4731eca77b9807cd708f298d47d34a75b46c728929e416dc47b0ee530ee8b058`.
- License: OfficeBench Apache-2.0; license and task record retained here.

## Independent truth and boundaries

The input has one `Tasks` sheet and twelve visible task names, `Activity 1` through `Activity 12`, in rows 3-14. The adapted solution converts the workbook to PDF with LibreOffice and attaches the resulting bytes to an `Alice.eml` draft. The upstream evaluator's `Activaty1` spelling is not source evidence; the local Judge checks the actual input spelling and all twelve names.

## Evidence-to-requirement map

| Requirement | Evidence | Judge |
|---|---|---|
| All task rows in PDF | input workbook task column | R002/R003 |
| PDF and email filenames | original query | R001/R004 |
| Alice recipient and review purpose | original query | R004/R005 |
| Attachment bytes equal generated PDF | adapted cross-artifact integrity rule | R005 |
| Input provenance | fixed source hash | R006 |

## Validation

Golden and fresh replay are generated with `solution/solve.sh` and checked with `tests/test.sh`. Negative cases include empty/corrupt PDF, missing attachment, wrong recipient, and a PDF missing one activity. `pdftotext` and Python MIME parsing provide deterministic checks; LibreOffice conversion is pinned in the environment.

## Difficulty, leakage, and open decisions

The task crosses spreadsheet rendering, PDF packaging, and email MIME creation with twelve source rows; author hypothesis is L3. Leakage group is `officebench-task-pdf-email`. Final difficulty and acceptance remain for human review.
