# Difficulty Design And Calibration

Read this before assigning `difficulty_tier` and again after rollout results exist.

## Difficulty is a hypothesis

Author labels are provisional. Record the proposed tier, structural evidence, expected failure modes, and what rollout result would falsify the label. Do not claim calibration until fixed Harness/model runs and human review are complete.

Use the repository mapping:

| Public level | `difficulty_tier` | Evidence-backed structure |
|---|---|---|
| L1 / Easy | `basic` | One application and one primary input; short, explicit edit/extraction; simple formulas or formatting; few independent checkpoints. |
| L2 / Medium | `intermediate` | One application with several stages or two inputs; common formulas, filters, pivoting, native charting, or moderate preservation; several result regions to verify. |
| L3 / Hard | `advanced` | Multiple files or sources, possibly cross-application; joins/retrieval plus analysis, formula linkage, charts or multiple deliverables; longer dependency chain or state-sensitive decisions. |
| L4 / Expert | `expert` | Long coherent professional workflow with at least three meaningful sources/applications or equivalent expert modeling depth; dynamic/conflicting state, multimodal evidence, error recovery, strict preservation, reusable processors, or many causally dependent checkpoints. |

## Structural rationale

Record evidence for these axes:

- number and heterogeneity of task-relevant inputs;
- applications and artifact types genuinely required;
- dependency-chain depth and order sensitivity;
- spreadsheet operations and preservation surface;
- state, time, conflict, and update handling;
- multimodal extraction or visual fidelity;
- professional/domain reasoning;
- number of independent checkpoints and realistic recovery burden;
- estimated competent-human completion time and its source.

Large row counts, many sheets, verbose instructions, foreign language, slow downloads, obscure filenames, or hidden requirements do not constitute difficulty.

## Designing toward a tier

Do not distort one task to fill a quota. Select a source whose native workflow already supports the desired tier, then preserve that structure.

- To obtain L1/L2, prefer direct single-workbook edits, extraction, formula completion, or native chart construction with explicit targets.
- To obtain L3, prefer coherent joins across files, policy application, audit/report packages, or offline reconstructions of a real multi-stage workflow.
- To obtain L4, prefer professional financial/model repair, dynamic multi-source reconciliation, conflicting updates, or processors that must generalize under mutation. Every extra source must affect the answer.

If an L3/L4 source cannot be made self-contained without deleting the state dependencies that make it hard, keep it as a workflow reference rather than downgrading it silently into a shallow spreadsheet task.

## Rollout calibration

After Golden, Judge, and negative calibration pass, use the frozen matrix with no solver Skills or authoring assistance. Keep Golden, Rubric, Judge, author records, and authoring utilities hidden. Codex and OpenCode may use only their built-in capabilities plus the task-visible environment and inputs.

Initial operating ranges for pass@1 are:

- L1: `>= 0.80`
- L2: `0.55-0.79`
- L3: `0.25-0.54`
- L4: `0.05-0.24`
- `< 0.05`: investigate ambiguity, broken environment, weak instructions, or infeasibility before calling it L4

Use pass@1 as the main signal and mean Judge score, pass@k, failure type, cost, and step count as supporting evidence. Tier medians should decrease monotonically, with a target adjacent gap of at least `0.15`.

Codex and OpenCode runs that also change the underlying model are cross-replication evidence, not a causal Harness comparison. When structural rationale and rollout disagree, inspect the task and Judge first, then relabel, revise, or reject with a recorded reason. A human approves the calibrated tier.
