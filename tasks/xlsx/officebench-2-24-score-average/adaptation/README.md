# C017 adaptation record

Status: `pending_human_review`. The `intermediate` tier is an L2 author hypothesis, not a calibrated release label.

## Source and evidence

- Source: OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`, `tasks/2-24/subtasks/0.json`.
- Original request: `Extract the score and name then calculate the average score in excel score.xlsx`.
- Input: `tasks/2-24/testbed/data/score.xlsx`, 6,389 bytes, SHA256 `e15533ba4ea8f4bcddf454c9e05d0f34fcc511a20041600c72b2a98ae9c3551d`.
- License: OfficeBench Apache-2.0; license and task record retained here.

## Independent truth and boundaries

The visible input has Bob=62, Henry=94, and Quinn=81. Their arithmetic mean is 79.0. The adapted task preserves those rows and appends one `Average` row. The upstream evaluator only searched for the keyword `79`; the local Judge checks the complete table and rejects missing, duplicate, or changed students.

## Evidence-to-requirement map

| Requirement | Evidence | Judge |
|---|---|---|
| Preserve names and scores | visible input workbook | R002/R004 |
| Average value | independent sum of input scores | R003 |
| Native editable workbook | adapted instruction and OOXML inspection | R001/R005 |
| Input provenance | fixed source hash | R006 |

## Validation

Golden and fresh replay are generated with `solution/solve.sh` and checked with `tests/test.sh`. Required negatives are empty/corrupt output, unchanged input without Average, wrong average, and a duplicated student. Repeated deterministic checks and LibreOffice/openpyxl open/render checks are required before handoff.

## Difficulty, leakage, and open decisions

This is a short single-workbook extraction plus aggregation task; author hypothesis is L2. Leakage group is `officebench-score-average-small`; only this input variant is in the group. Final difficulty and acceptance remain for human review.
