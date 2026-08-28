# Validation And Human Handoff

Read this after implementing a candidate. Run cheaper deterministic checks before Harbor builds or model rollout.

## Validation ladder

### 1. Static package checks

From the repository root:

```bash
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/xlsx/<task-id>
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/xlsx/<task-id>
```

Also validate JSON/TOML, executable entrypoints, input-manifest sizes/hashes, byte identity between `data/input_files/` and `environment/input_files/`, metadata consistency, Rubric/Judge criterion alignment, exact output-name agreement, and absence of placeholders or credentials.

### 2. Truth and solution checks

- Independently recompute key facts from inputs and rules.
- Run `solution/solve.sh` from a clean output directory.
- Judge the checked-in Golden and the fresh replay; both should satisfy every intended hurdle and normally score `1.0`.
- Compare the replay with source truth semantically. Byte equality is optional unless the format itself requires it.
- Recalculate formula workbooks and open/render them in the packaged runtime. Inspect every relevant sheet/page, not only the first.

### 3. Negative calibration

At minimum test:

- empty output directory;
- missing required file;
- corrupt or non-OOXML workbook;
- unchanged input renamed as output when the task is an edit;
- one core numeric/formula/business error;
- one omitted row/section or wrong join/mapping;
- collateral change outside the allowed edit region;
- screenshot or hardcoded-value substitute where native editability is required;
- cross-artifact inconsistency for multi-output tasks.

Create at least one **medium defect** that should receive partial credit below the pass threshold. If every non-Golden artifact scores zero or near one, the Judge has not demonstrated useful gradient.

For reusable processors or stateful formulas, add a metamorphic case that changes a source value, adds another instance of an error family, or toggles selections. The correct output must change accordingly while unrelated values remain preserved.

Repeat deterministic Judge runs to confirm identical results. For model/vision criteria, run enough repeats to measure variance and require the repository's standard-deviation bound before release.

### 4. Harbor checks

Run representative official Harbor trials only after local checks pass:

- Oracle/fresh solution positive;
- no-op or empty negative;
- a targeted medium defect when the Harbor interface permits it.

Record infrastructure failures separately from model/task failures. A trial that stops during image build or Agent installation is not a task score.

### 5. Difficulty rollout

Use the frozen rollout configuration and digest all relevant task, Harness, prompt, Skill, and Judge inputs. Record attempt-level artifact, score, success, failure category, steps, tokens, cost, and runtime. Do not tune hidden checks around one model's accidental behavior.

## Adaptation record

`adaptation/README.md` should state:

1. why this exact upstream task was selected;
2. immutable source paths/revisions, licenses, and input hashes;
3. every input and upstream artifact actually inspected;
4. facts independently derived from source inputs;
5. all translations, projections, side-effect replacements, corrections, removals, and added audit layers;
6. why the local Golden and Judge are justified when upstream evidence conflicts;
7. positive, negative, metamorphic, render, and Harbor results with commands;
8. initial difficulty axes and uncalibrated tier;
9. leakage group and known near-duplicates;
10. unresolved license, fidelity, environment, Judge, or ambiguity risks.

Do not write "validated" without naming which layer passed.

## Human review packet

End with status `pending_human_review` and provide reviewers:

- source and asset license decision points;
- original versus adapted instruction diff;
- input inventory and rendered previews;
- evidence-to-requirement/Rubric map;
- Golden facts and independent recomputation evidence;
- Judge score table for Golden, replay, empty, corrupt, unchanged, medium, and targeted negative cases;
- author difficulty hypothesis plus rollout evidence when available;
- leakage group and unresolved risks.

Codex may fix deterministic defects found during authoring, but must not fill reviewer names, reviewer decisions, final difficulty approval, or final release status. Human rejection returns the task to authoring with the stated issue; human acceptance is still separate from dataset split and release freezing.
