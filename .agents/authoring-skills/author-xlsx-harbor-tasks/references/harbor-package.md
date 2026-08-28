# CoworkEvoBench Harbor Package Contract

Read this while constructing the task package. The repository schemas and validators remain authoritative if they change.

## Required structure

```text
tasks/xlsx/<task-id>/
|-- instruction.md
|-- task.toml
|-- rubric.json
|-- environment/
|   |-- Dockerfile
|   `-- input_files/
|-- data/input_files/
|-- solution/
|   |-- solve.py
|   |-- solve.sh
|   `-- <reviewed Golden deliverables>
|-- tests/
|   |-- judge.py
|   |-- test.sh
|   `-- expectations.json
|-- metadata/
|   |-- task.json
|   |-- provenance.json
|   `-- input-manifest.json
`-- adaptation/
    |-- README.md
    `-- retained upstream evidence and licenses
```

Start from `tasks/templates/harbor-office-task/`, then compare against the closest existing Excel task. Remove every placeholder.

## Visibility and paths

The solver sees only the instruction, allowed public metadata, and containerized inputs. The trusted evaluation side owns Golden, Rubric, Judge, upstream audit material, and split metadata.

- Copy identical paths and bytes into `data/input_files/` and `environment/input_files/`.
- The Dockerfile copies `environment/input_files/` to `/app/data/input_files/`, creates `/app/output`, and makes the same inputs available to the verifier at `/data/input_files`.
- `task.toml` uses `artifacts = ["/app/output"]`, `allow_internet = false`, and `network_mode = "no-network"` for public offline tasks.
- `tests/test.sh [artifact_dir]` writes the structured result under `/logs/artifacts/` and scalar reward to `/logs/verifier/reward.txt`.
- Pin runtime dependencies and fonts that affect parsing, formula behavior, rendering, or pagination.

## Instruction

Write a standalone user request, not an evaluation checklist. Preserve the upstream intent and state:

- identify inputs and their roles;
- state exact deliverable filenames and `/app/output` paths;
- expose requirements the solver should reasonably know, including preservation and editability;
- make dates, units, sorting, joins, and business definitions explicit only when supported by source evidence;
- avoid leaking Rubric weights, hidden target cells, Golden-only facts, or implementation details.

Translation may clarify file paths and offline substitutions, but must not strengthen the task silently. Retain the original query under `adaptation/`.

## Metadata and provenance

Keep `task.toml [metadata]` and `metadata/task.json` aligned with current schemas and taxonomy. Use only the existing `difficulty_tier` field; record L1-L4 reasoning elsewhere rather than creating duplicate labels.

`metadata/provenance.json` records the immutable source URI/revision, license, privacy status, whether wording was rewritten, conservative `synthetic` status, publishability, and hashes of all solver-visible inputs. `metadata/input-manifest.json` records each input's local path, SHA256, size, and exact source or projection.

Do not label a task public while asset rights are pending. Use restricted/prohibited status or leave it in the candidate inventory.

## Golden and reference solution

- Build Golden from visible inputs, rules, and reviewed source truth.
- Preserve editable formulas, charts, controls, tables, styles, and document structure required by the task.
- Check calculations independently and inspect all rendered pages/sheets that matter.
- `solution/solve.py` must reproduce this one task from its inputs. It may contain task-specific reviewed facts when those facts are traceable, but it must not be presented as a general generator.
- When the task requires the solver to deliver a reusable processor, test that processor on a meaningful input mutation. The reference solution itself must also pass that test.

## Rubric

Use at least six atomic criteria and keep `rubric.json`, Judge criterion IDs, weights, types, descriptions, and pass threshold aligned.

- Put file existence/openability and truly indispensable outcomes behind hurdle caps.
- Give core data accuracy, formula logic, and business correctness more weight than filenames or cosmetic formatting.
- Add penalties for plausible harmful failures: unsupported claims, unchanged input, collateral damage, missing rows, hardcoded results, fake logs, screenshots replacing native objects, or cross-artifact conflicts.
- Do not double-count the same failure across many criteria.
- Prefer deterministic checks. Use semantic or visual judgment only where multiple valid outputs make deterministic checking insufficient.
- A criterion must be supported by the user-visible instruction or a universal professional integrity requirement; do not turn Golden quirks into hidden requirements.

## Judge

Recompute expected facts from `/data/input_files` or frozen expectation evidence. Compare semantics and invariants rather than requiring byte equality with Golden.

For XLSX, check the relevant subset of:

- OOXML/openability and exact requested outputs;
- sheet names/order, allowed change regions, formulas versus cached values, number formats, defined names, merged cells, validations, conditional formatting, charts/images, VBA preservation, print settings, and hidden state;
- independent totals, joins, unmatched records, unit/date conversion, boundary formulas, and cross-artifact consistency;
- LibreOffice recalculation and rendering with the packaged font environment;
- behavior after changed selections or source values;
- absence of unrelated modifications and unsupported facts.

Return per-criterion evidence, normalized reward, pass status, and hurdle failures. Missing or corrupt files must fail cleanly rather than crash the verifier.
