# officebench-2-21-bill-courses v0.3 adaptation record

Status: `harbor_replayed`; difficulty remains `unassigned`; `calibration_status=uncalibrated`.

## Upstream seed, evaluator, and license

- Dataset: OfficeBench task `2-21/0`, Git commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Frozen Query/Evaluator arguments: `adaptation/source-task.json`, SHA256 `44c5906595c29ecf33fcecb81728e814710e1eab70bf9c3fb3e6a0b0216eb07c`.
- Frozen input: upstream `tasks/2-21/testbed/data/Bill_course_lists.xlsx`, SHA256 `d7764b78e2931fbcda87f2a45ef140693c3d6977e0235b264aa868cf949bab8a`. The upstream, `data/input_files`, and `environment/input_files` copies are byte-identical.
- Evaluator implementation consulted: upstream `utils/evaluate.py`, SHA256 `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90`. Its `evaluate_contain` path case-insensitively searched extracted document text for 15 course-name substrings only.
- License: Apache-2.0. `adaptation/LICENSE-OfficeBench` is byte-identical to the upstream `LICENSE`, SHA256 `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`.

## Full context read before authoring

Every existing package file was read directly: instruction, TOML, rubric, Dockerfile, both input copies, manifests, source task/license, adaptation files, metadata, solver/wrapper, Judge/test wrapper/expectations, and both old DOCX outputs. The workbook was read twice, with formulas visible and with cached values, cell-by-cell with styles and properties. It contains one visible `Sheet1`, range `A1:D16`, no formulas, merged cells, hidden rows/columns, names, freeze panes, or filters. The old Golden was also parsed and rendered: one portrait Letter page containing 15 Normal-style course-name paragraphs and no tables, headers, footers, links, controls, or traceability.

The exact upstream `2-21/0` task JSON, all six related `2-21` subtasks, OfficeBench README/license/revision, source workbook, and the full relevant `evaluate_contain` implementation were inspected. All six upstream variants reduce to the same course-name list with the same 15 positive keywords, so the former task could not detect wrong terms, years, grades, source rows, calculations, record order, unsupported claims, or flattened Word content.

## Input truth and calculation ledger

The 15 source rows are:

| Row | Term | Year | Course | Grade | Points |
|----:|------|-----:|--------|-------|-------:|
| 2 | Winter | 2022 | CS 101 | A | 4.0 |
| 3 | Winter | 2022 | MATH 150 | B+ | 3.3 |
| 4 | Winter | 2022 | PHYS 101 | A- | 3.7 |
| 5 | Spring | 2022 | CS 102 | A | 4.0 |
| 6 | Spring | 2022 | MATH 151 | B | 3.0 |
| 7 | Spring | 2022 | PHYS 102 | A- | 3.7 |
| 8 | Fall | 2022 | CS 103 | B+ | 3.3 |
| 9 | Fall | 2022 | MATH 152 | A | 4.0 |
| 10 | Fall | 2022 | PHYS 103 | B | 3.0 |
| 11 | Winter | 2023 | CS 104 | A- | 3.7 |
| 12 | Winter | 2023 | MATH 250 | A | 4.0 |
| 13 | Winter | 2023 | PHYS 104 | B+ | 3.3 |
| 14 | Spring | 2023 | CS 105 | A | 4.0 |
| 15 | Spring | 2023 | MATH 251 | B+ | 3.3 |
| 16 | Spring | 2023 | PHYS 105 | A- | 3.7 |

The grade-point mapping is an explicit v0.3 task rule, not an inferred upstream fact. Term point totals/GPA in first-appearance order are `11.00/3.67`, `10.70/3.57`, `10.30/3.43`, `11.00/3.67`, and `11.00/3.67`; cumulative point/GPA endpoints are `54.00/3.60`. Subject totals/GPA are CS `19.00/3.80`, MATH `17.60/3.52`, and PHYS `17.40/3.48`. Grade counts A/A-/B+/B are `5/4/4/2`. Across twelve within-subject adjacent code edges, eleven are `+1`; MATH 152 to 250 is the sole `JUMP +98`. Exact duplicate rows and blank A:D cells are both zero.

## Adaptation boundary and leakage control

Version 0.3 supersedes the one-page course-name list in v0.2. Apache-2.0 permits the derivative adaptation. OfficeBench contributes the workbook facts, original extraction intent, Query/evaluator record, and license. CoworkEvoBench contributes the explicit unweighted grade-point policy, evidence-locator workflow, term/cumulative and subject reconciliation, numeric code-spacing audit, closure controls, native Word review controls, dedicated Oracle, rubric, and directed failures.

The task does not invent credits, prerequisites, degree requirements, transfer status, graduation status, or institutional standing. It remains in leakage group `officebench-bill-course-list`; all six upstream 2-21 paraphrases are treated as one source cluster. This task's academic-record reconciliation and editable term-review controls are deliberately different from OCR tasks `officebench-2-6-homework-ocr` and `officebench-3-22-homework-folder` and from the CS161 roster cluster.

## Golden and scoring boundary

The only deliverable is `bill_course_history_audit.docx`. The task-specific solver constructs editable text, nine native tables, five unlocked tagged rich-text content controls, repeating table headers, non-splitting rows, page header, and automatic page numbers. The Judge reads neither Golden nor solver-only markers: it independently rebuilds every row, source locator, mapped point, term and cumulative metric, subject sequence, continuity edge, distribution, and closure control from the frozen XLSX and the public mapping. It also parses OOXML safety/native structures and renders with LibreOffice.

Golden and fresh replay must repeatedly score 1.0. Empty, corrupt, flattened, extra-output, chronology, mapping, source-locator, term, cumulative, subject, continuity, distribution, native-control, locked-control, unsupported-claim, and closure failures must fail. The deliberately incomplete closure sample must retain stable nonzero partial credit. Oracle success is not real-Agent difficulty calibration; difficulty remains unassigned.

## Local validation and visual QA

`validate-task`, `validate-task-assets`, package `compileall`, JSON parsing, input-copy hash checks, ownership checks, and `git diff --check` pass. Golden and an independently generated fresh replay are byte-identical, SHA256 `5feea2b880fdad675994e49dafefb1696baffafd3226655b0f2868a8a22fac60`. Five Golden and five fresh Judge runs each score `1.0`; all ten result JSON files are byte-identical, SHA256 `8c977d0c797e93061f3b7090f5944d61f55fd5b8370379d6bc60b86cf645da18`.

Nineteen non-Golden cases all fail: empty, corrupt, flattened, and unexpected extra DOCX are `0.0`; missing record/wrong source order are `0.35`; wrong point mapping/locator are `0.52`; wrong term/cumulative values are `0.65`; wrong subject total/sequence are `0.72`; wrong continuity delta/locator are `0.76`; wrong distribution and the medium missing-one-Closure-row sample are `0.78`; missing/locked native controls and the unsupported graduation claim are `0.80`. The medium sample is strictly between zero and Golden.

LibreOffice 7.3.7.2 renders Golden and fresh as four portrait Letter pages. Golden/fresh, six representative native negatives, the two-page flattened case, and Golden/fresh DOCX round-trips account for 50 checked pages. The Golden pages and every changed representative page were inspected at original render resolution; unaffected negative pages were proven pixel-byte-identical to the corresponding Golden render. There are no unintended blank pages, clipped or overlapping tables, off-page content, split data rows, unreadable fonts, or orphaned headings. The flattened case is visibly a title-only page followed by a black raster page.

LibreOffice can open, save, and export both valid documents without losing any visible content: the round-trips remain four readable pages and preserve all nine tables, five editable review-control values, and every business fact. LibreOffice 7.3 clears the Word-specific `w:tag` values on those content controls; therefore the strict Judge correctly scores each round-trip `0.80` instead of hiding this interoperability loss. The original DOCX and normal task replay retain all five stable tags and score `1.0`.

Official Harbor 0.21.0 Docker Oracle A/B ran against the same frozen task snapshot with `n_concurrent_trials=1`, `force_build=true`, full verifier, and no Agent network dependency. A job `docx-hardening-b8-bill-course-a-20260908` / trial `officebench-2-21-bill-courses__mfgfvqB` and B job `docx-hardening-b8-bill-course-b-20260908` / trial `officebench-2-21-bill-courses__PrBUNPM` each scored `1.0` with zero exceptions. Both report task checksum `6832c63c30d59b4d9a62f0edb248f6f121fb74c84db0e88bf4de718fd2bfc5a0` and digest `sha256:c75b1e86f59c5f9ce28c23d827902e0293c62175e18e86b80f4cfc434eb67bcf`. Their DOCX SHA256 is `5feea2b880fdad675994e49dafefb1696baffafd3226655b0f2868a8a22fac60`, Judge JSON SHA256 is `8c977d0c797e93061f3b7090f5944d61f55fd5b8370379d6bc60b86cf645da18`, and artifact manifest SHA256 is `3b2200d91f247c9edd3849a8c80e83402294fbabaf72c2f7e2523134acac9c2c` in both runs.
