# officebench-2-20-cs-major-list v0.3 adaptation record

Status: `local_validated`; difficulty remains `unassigned`; `calibration_status=uncalibrated`.

## Upstream seed, evaluator, and license

- Dataset: OfficeBench task `2-20/1`, Git commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Frozen Query/Evaluator arguments: `adaptation/source-task.json`, SHA256 `9dcb7ebe091e7e470ec408dd706f0d9fb97c9526f4feaaa9a3b9dba01ad5c755`.
- Frozen input: upstream `tasks/2-20/testbed/data/CS161_Class_Roster.xlsx`, SHA256 `c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7`. The upstream, `data/input_files`, and `environment/input_files` copies are byte-identical.
- Evaluator implementation consulted: upstream `utils/evaluate.py`, SHA256 `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90`; the task-relevant containment semantics are frozen in `adaptation/upstream-evaluator.py`.
- License: Apache-2.0. `adaptation/LICENSE-OfficeBench` is byte-identical to the upstream `LICENSE`, SHA256 `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`.

## Full context read before authoring

The author inspected every existing package file: instruction, task TOML, rubric, Dockerfile, both input copies, input manifest, source task, license, adaptation notes, provenance/task metadata, solver and shell wrapper, Judge/test wrapper/expectations, both old DOCX copies, and the old Golden's OOXML and one-page LibreOffice render. The workbook was read cell-by-cell, including styles and workbook properties: one worksheet named `Sheet1`, 10 rows × 10 columns, no formulas, merged cells, hidden dimensions, defined names, freeze panes, or filters. The old Golden contained only three Normal-style paragraphs and no tables, comments, links, bookmarks, headers, footers, or evidence locators.

The matching upstream `2-20/1` JSON, all three related `2-20` subtasks, exact source workbook, OfficeBench README, license, Git revision, and `evaluate_contain` implementation were also inspected directly. Upstream containment checked only nine keywords and could not validate filtering, order, score evidence, calculations, comments, navigation, editability, or privacy boundaries.

## Input truth ledger

Exactly three rows have `Major = CS`:

| Source row | Student | Standing | HW1 | HW2 | Midterm | HW3 | Final | Avg | Low | High | Range | M->F | Tier |
|------------|---------|----------|----:|----:|--------:|----:|------:|----:|-----|------|------:|-----:|------|
| `Sheet1!2` | `S1000 Zoye Pivozu` | Junior | 22 | 61 | 68 | 74 | 1 | 45.2 | Final (1) | HW3 (74) | 73 | -67 | Priority 1 |
| `Sheet1!3` | `S1001 Sipaq Cixarim` | Freshman | 58 | 77 | 88 | 18 | 19 | 52.0 | HW3 (18) | Midterm (88) | 70 | -69 | Priority 2 |
| `Sheet1!7` | `S1005 Tukemoy Yehamucu` | Freshman | 44 | 54 | 65 | 13 | 2 | 35.6 | Final (2) | Midterm (65) | 63 | -63 | Priority 1 |

The deterministic review order is `S1000 -> S1005 -> S1001`. Assessment-level aggregate truth is:

| Assessment | CS N | CS avg | Non-CS N | Non-CS avg | CS - non-CS |
|------------|-----:|-------:|---------:|-----------:|------------:|
| HW1        | 3 | 41.3 | 6 | 25.0 | +16.3 |
| HW2        | 3 | 64.0 | 6 | 55.2 | +8.8 |
| Midterm    | 3 | 73.7 | 6 | 43.3 | +30.3 |
| HW3        | 3 | 35.0 | 6 | 58.2 | -23.2 |
| Final      | 3 | 7.3 | 6 | 77.2 | -69.8 |

## Adaptation boundary and task differentiation

Version 0.3 supersedes the three-line `cs_class_2024.docx` extraction in v0.2. Apache-2.0 permits this derivative adaptation. The retained source contributions are the workbook facts, the exact `Major = CS` selection intent, source Query, and evaluator record. CoworkEvoBench contributes the review-tier rules, evidence-record structure, calculation contract, aggregate comparison, native comments, bookmarks/internal links, document-control requirements, input-driven Oracle, partial-credit rubric, and negative-case suite.

This task intentionally differs from sibling `officebench-2-20-all-cs161`: that sibling is a full-class source-order roster control book, while this task suppresses all non-CS identities, changes to a tier-sorted three-student review queue, creates one evidence record per selected student, requires anchored Word comments, and compares CS versus non-CS only through aggregates. It also differs from `officebench-2-20-final-perfect`, which remains reserved for Final=100 exception evidence. All stay in leakage group `officebench-cs161-roster` and must not be split across benchmark partitions.

## Golden, Oracle, and scoring boundary

The only deliverable is `cs_major_advising_docket.docx`; the source XLSX is input, not output. The task-specific solver constructs editable text/tables, three comments, three bookmarks/links, page header, automatic page number, repeating headers, and non-splitting rows. The Judge never reads Golden bytes or solver markers. It independently recomputes the CS selection, queue order, tier, deltas, per-student measures, source locators, cohort aggregates, control counts, and privacy exclusions from the frozen XLSX. It parses native OOXML for comments/bookmarks/links and safety properties, then performs a LibreOffice PDF render audit.

Golden and fresh replay must score 1.0 repeatedly; empty/corrupt/flattened outputs and task-specific filter, order, rule, evidence, aggregate, navigation, comment, and privacy failures must be rejected. A deliberately incomplete but otherwise correct file must receive a stable partial score between zero and Golden. Oracle success is not evidence of real-Agent difficulty; difficulty remains unassigned until a separate multi-Harness calibration.

## Visual QA record

LibreOffice 7.3.7.2 renders the Golden and fresh replay as five portrait Letter pages: review queue/rule card, three separately paged evidence records, and aggregate/control appendix. Golden, fresh, wrong queue order, wrong profile score, non-CS identity leak, medium missing-control, and flattened-image outputs were exported and all 32 pages were inspected at original render resolution. The valid files have no blank pages, clipping, overlap, off-page content, split tables, missing fonts, or unreadable text. Directed defects remain visibly localized; the flattened case visibly produces a mostly blank title page plus a black full-page raster. Golden and fresh LibreOffice DOCX round-trips each retain a Judge reward of 1.0, including comments, bookmarks, and links.

Local validation on 2026-09-07: `validate-task`, `validate-task-assets`, Python compilation, JSON parsing, and `git diff --check` pass. Golden and fresh replay each score 1.0 in five consecutive runs; all ten Judge JSON payloads are byte-identical with SHA256 `ffe6b122d04b83d7da3897041b193ea8d8ea4a8ef8c3842b62c923fa23910ff7`. Golden, fresh, and checked-in solution are byte-identical with SHA256 `e742f6a8cb777993d3a0c1864b5b03ea97b2011590df6c2118d8a38cf50a4259`.

Negative calibration scores are: empty `0.0`, corrupt `0.0`, flattened screenshot `0.0`, wrong filter/foreign student `0.35`, non-CS identity leak `0.35`, wrong queue order `0.62`, wrong tier boundary `0.62`, wrong profile score `0.68`, wrong source locator `0.68`, missing HW3 evidence `0.68`, wrong student average `0.78`, wrong cohort difference `0.78`, missing internal links `0.78`, and missing native comments `0.78`. The medium sample omits only one Control Summary row and scores `0.99`, strictly between zero and Golden while retaining all core evidence.
