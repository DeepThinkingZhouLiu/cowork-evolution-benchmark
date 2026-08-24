---
license: apache-2.0
language:
- zh
- en
pretty_name: OmegaUse-OfficeVal
size_categories:
- n<1K
task_categories:
- other
tags:
- office
- agents
- llm-agents
- benchmark
- computer-use
- economic-grounding
- document-processing
- spreadsheet
- presentation
configs:
- config_name: default
  default: true
  data_files:
  - split: train
    path: data-studio/train.parquet
dataset_info:
  features:
  - name: id
    dtype: string
  - name: instruction
    dtype: string
  - name: operation_intent
    dtype: string
  - name: domain
    dtype: string
  - name: human_labor_time
    dtype: int64
  - name: task_price_proxy
    dtype: int64
  - name: price_source
    dtype: string
  - name: origin_files_json
    dtype: string
  - name: rubrics_json
    dtype: string
  splits:
  - name: train
    num_bytes: 653586
    num_examples: 100
  download_size: 275125
  dataset_size: 653586
---

# OmegaUse-OfficeVal

**Benchmarking LLM Agents on Long-Horizon Office-Suite Tasks with Economic Grounding**

OmegaUse-OfficeVal is a benchmark for evaluating LLM agents on long-horizon,
real-world office-suite tasks that span word-processing documents, spreadsheets,
presentations, and cross-file productivity workflows. Tasks are derived from
authentic office requests proposed by practitioners and drawn from freelance
platforms, grounding the benchmark in real economic demand. Each task provides a
high-level user instruction and input artifacts, and requires the agent to
produce a final deliverable such as a Word document, spreadsheet, or presentation.
Evaluation focuses on the quality and correctness of the final deliverable rather
than on the specific execution trajectory.

A distinctive feature of OmegaUse-OfficeVal is its **economic grounding**: every
task is annotated with two complementary signals — *human labor time* and a
*task price proxy* — so that agent performance can be analyzed not only by task
completion, but also in terms of the human effort and economic value associated
with each task.

## Dataset Summary

| Property | Value |
| --- | --- |
| Domain | Office-suite productivity: Word, PowerPoint, Excel, PDF, and cross-file workflows |
| Tasks | 100 long-horizon tasks |
| Input files | 220 files (2.2 per task on average) |
| Golden artifacts | 115 files |
| Modalities | Text instructions, Office documents, PDFs, images, audio, and video |
| Evaluation | Deterministic, code-based verifiers over final delivered artifacts |
| Languages | Chinese and English task definitions and rubrics |
| Default Data Studio data | English task definitions merged with English rubrics |

## Repository Layout

This Hugging Face dataset repository contains the input artifacts, bilingual task
definitions, bilingual evaluation rubrics, and a merged English file for Data
Studio:

```
task_files/
  officeval_001/
    01_....docx
    02_....docx
  ...
  officeval_100/

tasks/
  officeval_001.json
  ...
task-en/
  officeval_001.json
  ...
rubrics/
  officeval_001.json
  ...
rubrics-en/
  officeval_001.json
  ...

data/
  train-00000-of-00001.parquet

data-studio/
  train.jsonl
  train.parquet

tasks_and_rubrics_en.json
tasks_and_rubrics_en.parquet
```

| Repository path | Content | File/record count |
| --- | --- | ---: |
| `task_files/officeval_<NNN>/` | Input artifacts grouped by task | 220 files across 100 folders |
| `tasks/` | Chinese task definitions | 100 JSON files |
| `task-en/` | English task definitions | 100 JSON files |
| `rubrics/` | Chinese evaluation rubrics | 100 JSON files |
| `rubrics-en/` | English evaluation rubrics | 100 JSON files |
| `data-studio/train.parquet` | Page-indexed, Viewer-optimized English Data Studio split | 100 rows |
| `data-studio/train.jsonl` | Flat JSONL copy of the English Data Studio split | 100 rows |
| `data/train-00000-of-00001.parquet` | Structured Parquet copy of the English split | 100 rows |
| `tasks_and_rubrics_en.json` | Merged English source data | 100 records |
| `tasks_and_rubrics_en.parquet` | Parquet copy of the merged English source data | 100 rows |

Input files can be downloaded directly with the following URL pattern:

```
https://huggingface.co/datasets/baidu-frontier-research/OmegaUse-OfficeVal/resolve/main/task_files/officeval_<NNN>/<filename>
```

The same task definitions and evaluation rubrics are also released alongside the
evaluation harness in the companion GitHub repository:

```
https://github.com/baidu-frontier-research/OmegaUse-OfficeVal
```

In this repository, `tasks/officeval_<NNN>.json` holds each task's Chinese
instruction, and `task-en/officeval_<NNN>.json` holds its English translation.
Both files also include the task's
`operation_intent`, `domain`, `human_labor_time`, `task_price_proxy`,
`price_source`, and the list of `origin_files` (pointing back to the files in this
Hugging Face repository). The corresponding Chinese and English scoring rubrics
are stored in `rubrics/officeval_<NNN>.json` and
`rubrics-en/officeval_<NNN>.json`, respectively.

## Data Studio

The default Data Studio view uses `data-studio/train.parquet`. It is generated from
`tasks_and_rubrics_en.json`, and each row combines an English task definition
with its matching English rubric. To maximize Viewer compatibility, nested
`origin_files` and `rubrics` values are serialized losslessly as JSON strings;
the file uses small Parquet row groups and a page index. Structured and JSONL
copies remain available in the repository.

| Data Studio property | Value |
| --- | --- |
| Configuration | `default` |
| Split | `train` |
| Rows | 100 |
| Source language | English |
| Data file | `data-studio/train.parquet` |

| Column | Type | Description |
| --- | --- | --- |
| `id` | string | Stable task identifier from `officeval_001` to `officeval_100` |
| `instruction` | string | English task instruction |
| `operation_intent` | string | Intended office operation |
| `domain` | string | Task domain |
| `human_labor_time` | integer | Human completion time in minutes |
| `task_price_proxy` | integer | Estimated task value in yuan |
| `price_source` | string | Source of the task price proxy |
| `origin_files_json` | string | JSON-serialized input artifact URLs and destination filenames |
| `rubrics_json` | string | JSON-serialized `dim1` usability and `dim2` task-completion items |

Browse and filter the records in the
[OmegaUse-OfficeVal Data Studio](https://huggingface.co/datasets/baidu-frontier-research/OmegaUse-OfficeVal/viewer).
The dataset can also be loaded directly with the Hugging Face `datasets` library:

```python
from datasets import load_dataset

dataset = load_dataset("baidu-frontier-research/OmegaUse-OfficeVal")
print(dataset)
print(dataset["train"][0])
```

## Task Components

Following the paper, each task consists of five components:

- **Instruction** — a textual description of the user request and the specific
  requirements to be satisfied.
- **Input files** — the collection of files required to complete the task
  (hosted in this repository).
- **Human labor time** — the recorded time required by human workers to complete
  the task *without* LLM assistance (in the released task metadata this is stored
  in **minutes**).
- **Task price proxy** — a task-level price signal estimating the market price of
  completing the task (in the released task metadata this is stored in **yuan**).
- **Code verifier** — executable evaluation code derived from the instruction and
  rubric that assigns a score to a candidate artifact.

## Economic Grounding

Two complementary indicators ground each task in economic reality:

- **Human labor time.** Each task is completed by at least two recruited
  annotators under a quality-gated protocol; annotators first pass interviews and
  sample tasks. When two valid completion times differ substantially, a third
  annotator is added. The reported human labor time is the average of the two
  shortest valid completion times, reducing the influence of unusually slow
  attempts.
- **Task price proxy.** A hybrid strategy is used. Tasks with **explicit price
  signals** provided by practitioners use those values directly
  (`price_source = explicit_price`). For tasks without explicit prices, three
  domain experts independently estimate the value and their estimates are
  aggregated through a consistency-based procedure that discards clear outliers
  (`price_source = estimated_price`).

## Task Distribution

Based on the required output format, tasks are categorized into five types:

| Required output type | Tasks | Share |
| --- | ---: | ---: |
| Word | 37 | 37% |
| PowerPoint | 39 | 39% |
| Excel | 19 | 19% |
| PDF | 4 | 4% |
| Cross-file (Word + Excel) | 1 | 1% |
| **Total** | **100** | **100%** |

Distribution of input files and golden artifacts (from the paper):

| File type | Input files | Golden artifacts |
| --- | ---: | ---: |
| Video/Audio | 10 | 0 |
| Image | 77 | 0 |
| DOCX | 63 | 48 |
| PPTX | 31 | 40 |
| XLSX | 25 | 24 |
| PDF | 14 | 3 |
| **Total** | **220** | **115** |

## Evaluation Protocol

OmegaUse-OfficeVal is evaluated with deterministic, code-based verifiers on the
final output files rather than by enforcing a fixed execution trajectory. Agents
may complete tasks through GUI actions, scripts, APIs, or hybrid strategies,
reflecting real office work where multiple valid workflows can lead to the same
deliverable.

Each task rubric has two dimensions:

- **Usability rubric (Dim-1).** Checks basic usability — correct file format, the
  file opens normally, content/layout is not severely corrupted, and the artifact
  remains editable. If any usability item fails, the task scores zero and the
  completion rubric is not evaluated further.
- **Task-completion rubric (Dim-2).** Fine-grained, user-centered scoring points
  covering content, formatting, structure, numerical accuracy, layout, and
  instruction following. Positive items credit correctly completed requirements
  (discrete weights of +1 / +3 / +5); negative items penalize unintended changes
  or avoidable damage that raise the user's downstream repair cost (symmetric
  penalties of −1 / −3 / −5).

The score is computed in two stages: an artifact must first pass **all** usability
items; its normalized task-completion score is then computed from the weighted
positive and negative rubric items, with the raw score clipped at zero so that
task completion is never negative.

| Rubric dimension | Total items | Average per task |
| --- | ---: | ---: |
| Usability checks (Dim-1) | 219 | 2.19 |
| Task-completion items (Dim-2) | 2009 | 20.09 |

## Data Construction & Privacy

The dataset is built through a multi-stage task-adaptation pipeline:

1. **Task collection.** Practitioners propose realistic office-suite tasks grounded
   in their everyday workflows, together with representative sample materials.
2. **Funnel screening.** An initial pool of practitioner-proposed tasks is
   progressively filtered down to a curated set, keeping only tasks that are
   grounded in real needs, have clear deliverables, and are sufficiently
   non-trivial and long-horizon.
3. **Privacy-preserving adaptation.** Instructions are rewritten to remove
   sensitive or identifying information while preserving the original user intent,
   constraints, and task-critical details. Subjective, unverifiable requirements
   (e.g., "make it beautiful") are removed, while natural, colloquial phrasing is
   retained.
4. **Input reconstruction & de-identification.** Input files are reconstructed with
   LLM assistance and manually revised to remove privacy/copyright risks, fix
   layout and consistency issues, and keep artifacts faithful to the intended task.
5. **Rubric generation & code verifiers.** Rubrics are generated with LLM
   assistance and iteratively refined by experts, then converted into executable
   code verifiers, with human–code discrepancy resolution to align automated and
   expert judgments.
6. **Final acceptance review.** Three senior experts confirm that each task poses
   no privacy risk, can be executed normally, and matches the language, structure,
   and artifact format of real office scenarios; a task is included only when all
   three experts agree.

| Funnel stage | Retained tasks |
| --- | ---: |
| Initial pool | 1,715 |
| Screening stage 1 | 595 |
| Screening stage 2 | 282 |
| Final benchmark | 100 |

## Intended Use & Limitations

OmegaUse-OfficeVal is intended as a reproducible, economically grounded testbed for
measuring progress toward reliable LLM agents for everyday office productivity. It
is **not** intended to measure whether LLM agents can replace entire occupations;
rather, it evaluates a practical setting in which a worker is assisted by a capable
agent. Because negative failure modes cannot be exhaustively enumerated, scores are
lower-bounded at zero and should be interpreted as relative measures of deliverable
quality.

## Citation

```bibtex
@inproceedings{omegause_officeval_2026,
  title     = {OmegaUse-OfficeVal: Benchmarking LLM Agents on Long-Horizon Office-Suite Tasks with Economic Grounding},
  booktitle = {tech report},
  year      = {July, 2026}
}
```
