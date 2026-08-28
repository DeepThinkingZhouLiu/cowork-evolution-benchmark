# Patterns From The Existing 14 Excel Tasks

Read this when choosing an adaptation pattern or checking whether a new task repeats a known failure.

## What the current tasks demonstrate

| Pattern | Existing examples | Reusable lesson |
|---|---|---|
| Correct a flawed upstream Golden | customer satisfaction; JSON cleaning | Recompute from inputs. Remove invented sample sizes, correlations, translations, or business impact. Preserve the upstream defect as evidence. |
| Project database/API state offline | inventory forecast; MOEX dashboard; regional forecast; SLA audit | Extract the minimum task-relevant state from the same fixture, retain projection provenance, freeze time, and exclude unrelated personal fields. |
| Replace external side effects | inventory calendar; pipeline Google Sheet; forecast Google Sheet; SLA email | Replace with a local schedule sheet, workbook, or message draft that preserves the information and remains deterministic. |
| Preserve an upstream task almost directly | OfficeVal 082/086/087/090/091/092 | Keep the original intent and real files; add output paths, preservation bounds, editability, recalculation, and Harbor packaging without inventing new business requirements. |
| Add an auditable derived layer | FORTE Finance-018 | A processor, integrated workbook, or trace log may be added when it exposes the original computation and is validated metamorphically; it must not change the upstream answer. |
| Repair a professional model | SpreadsheetBench-2 09_01 | Require formula reasoning, non-target preservation, replayable repair, mutation handling, native formulas, and structure/style regression checks. |

## Established authoring decisions

- An upstream task may be authentic while its fixture remains synthetic. Record both facts accurately.
- A task named "dashboard" does not justify hidden chart or visual requirements when the query asks only for tables.
- Upstream visual weakness is not automatically a solver failure. Score only requested or professionally necessary presentation properties that are visible in the instruction.
- Open-ended deliverables should be checked for supported content and consistency, not exact prose or a single person's name.
- Spreadsheet edits need preservation checks outside the allowed change region: sheets/order, formulas, constants, merges, styles, dimensions, names, validation, charts, images, print settings, and hidden state as relevant.
- Native editability matters. Reject screenshots standing in for requested tables, formulas, controls, or charts.
- Formula and stateful tasks need behavioral replay, not only static cell inspection. Change selections or source values, recalculate, and check the resulting behavior.
- Reusable processors need metamorphic tests. A script that copies a frozen Golden or branches on the canonical input hash is not a valid solution.
- Render requirements need the actual runtime, fonts, and pagination. Pin fonts or Fontconfig aliases when fidelity depends on them.

## Anti-patterns found or explicitly prevented

- Building a task from a taxonomy/category name without a source query and inputs.
- Treating the repository's root license as blanket permission for all attachments.
- Using an upstream Golden to fabricate missing market, customer, or financial inputs.
- Silently accepting unsupported facts because they appear in a reference answer.
- Translating or normalizing fields the original request did not authorize.
- Requiring visual polish absent from the user request to manipulate difficulty.
- Judging by workbook hash, exact internal implementation, or exact prose when multiple valid solutions exist.
- Accepting an unchanged input, empty template, hardcoded values instead of formulas, fixed-output processors, or collateral edits.
- Calling a task L4 because it has many rows or many formula cells.

## Selecting a closest exemplar

Choose the existing task with the closest *adaptation risk*, not merely the closest industry:

- live state to offline snapshot: use a Cowork-Bench static projection example;
- native chart or screenshot-guided edit: use OfficeVal 086 or 092;
- multi-workbook reconciliation and executable processor: use FORTE Finance-018;
- formula repair and mutation testing: use SpreadsheetBench-2 09_01;
- large preservation-sensitive template edit: use OfficeVal 087, 090, or 091;
- multi-artifact analytical consistency: use customer satisfaction or regional forecast.

Copy structure and checking strategy, never task-specific facts, coordinates, formula answers, or hidden expectations.
