---
name: author-xlsx-harbor-tasks
description: Audit public repositories and adapt evidence-backed spreadsheet tasks, inputs, and truth into review-ready CoworkEvoBench Harbor task packages. Use when sourcing, reconstructing, combining, grading, or validating Excel benchmark tasks; do not use to solve benchmark tasks or invent synthetic task stories from a taxonomy label.
metadata:
  standard: agent-skills
  domain: benchmark-authoring
  benchmark: cowork-evolution-benchmark
---

# Author Excel Harbor Tasks

Produce a candidate task that a human reviewer can accept, revise, or reject. Treat public repositories as an evidence environment: extract real task intent, inputs, rules, state, and truth from them. Do not treat repository names or task categories as prompts for inventing a similar business story.

The output of this Skill is never a released benchmark item. Stop at `pending_human_review`; a human makes the final task, license, difficulty, and release decision.

## Non-negotiable boundaries

- Preserve an auditable chain from every task requirement and scored fact to the upstream query, input, rule, state snapshot, or independently derived truth.
- Never infer input facts from an upstream Golden. Derive truth from visible inputs and rules; use the Golden only as evidence to audit.
- Do not copy assets whose redistribution rights are unclear. A repository license does not automatically cover every workbook, image, email, attachment, or linked file.
- Do not increase difficulty by adding rows, vague requirements, hidden expectations, or unrelated applications. Difficulty must arise from a coherent real workflow.
- Do not expose `solution/`, Golden files, Rubric details, Judge code, author notes, or split membership to the solver.
- Keep Judge and Golden logically independent. The Judge should recompute facts from inputs or frozen evidence rather than accept a reference file hash or one exact layout.
- Do not mark a task accepted, calibrated, or release-ready on Codex's own authority.

## Workflow

1. Read the repository's `AGENTS.md`, `docs/task-spec.md`, `tasks/README.md`, schemas, validator code, and the closest existing Excel adaptations before editing.
2. Audit the proposed source before authoring. Read [source-audit.md](references/source-audit.md) and create an evidence ledger. If the task, inputs, truth basis, or rights are missing, keep it as `source_screened`, `workflow_screened`, or `reference_only`; do not manufacture a task package.
3. Choose an adaptation mode:
   - **Direct asset adaptation:** preserve an upstream query and usable files; repair only documented defects and Harbor incompatibilities.
   - **Offline workflow reconstruction:** project the minimum task-relevant state from the same licensed fixture and replace external side effects with equivalent local, scoreable artifacts.
   - **Evidence-backed composition:** combine complementary sources only when they form one natural user workflow and each source has a declared role and compatible rights. Never splice unrelated tasks merely to raise difficulty.
4. Inspect every retained input, upstream evaluator, Golden, and rule. Record disagreements and decide them using the user request plus source evidence, not deference to the upstream Golden. Read [existing-task-patterns.md](references/existing-task-patterns.md) for patterns demonstrated by the 14 current Excel tasks.
5. Write an evidence-to-requirement map before implementation. Every instruction requirement and Rubric criterion must identify its evidence source and verification method. Open-ended professional judgment may allow multiple valid outputs; do not overfit it to Golden wording.
6. Assign an **author hypothesis**, not a final difficulty label. Read [difficulty-design.md](references/difficulty-design.md), record the structural rationale, and design only evidence-supported complexity.
7. Build the task from `tasks/templates/harbor-office-task/` and the closest accepted task. Read [harbor-package.md](references/harbor-package.md). Keep source preservation, task semantics, environment packaging, Golden construction, Rubric, and Judge as separate decisions.
8. Use repository-native tools and task-specific code for workbook inspection, editing, recalculation, rendering, and delivery QA. Keep all authoring instructions and utilities outside the solver workspace; difficulty rollouts receive no Skill or authoring assistance.
9. Validate in increasing-cost order. Read [validation-and-handoff.md](references/validation-and-handoff.md). At minimum require static validation, input/hash agreement, Golden, fresh replay, empty output, corrupt output, unchanged input where applicable, targeted medium defects, and render/recalculation checks.
10. Update `adaptation/README.md` with evidence, decisions, commands, results, known limitations, initial difficulty rationale, and unresolved risks. Hand the task to a human with status `pending_human_review` and do not continue to dataset splitting or release.

## Required authoring outputs

A completed candidate includes the Harbor task package plus:

- exact upstream repository, immutable commit or dataset revision, task path, acquisition date, license evidence, and per-input SHA256;
- original query/evaluator/rubric/Golden artifacts retained under `adaptation/` when redistribution permits;
- a clear record of original, projected, translated, corrected, removed, and newly required elements;
- an evidence ledger covering inputs, truth, transformations, instruction requirements, and Rubric criteria;
- an initial L1-L4 rationale and explicit `uncalibrated` status;
- validation evidence and a concise human-review packet listing all remaining judgment calls.

If any required evidence cannot be established, report the exact gap and stop before creating a misleading complete-looking package.
