---
license: cc-by-sa-4.0
---
# SpreadsheetBench 2 Dataset

SpreadsheetBench 2 evaluates spreadsheet agents on end-to-end business and financial workflows. The release contains 321 tasks in four categories: Debugging, Financial Model, Template, and Visualization.

## Overview

| Directory | Task type | Tasks | Input workbooks | Gold workbooks |
|---|---|---:|---:|---:|
| `Debugging` | Spreadsheet debugging and error correction | 100 | 100 | 10 |
| `Financial_Model` | Completion of multi-sheet financial models | 100 | 100 | 20 |
| `Template` | Completion of spreadsheet templates | 97 | 97 | 97 |
| `Visualization` | Chart creation and manipulation | 24 | 24 | 24 |
| **Total** |  | **321** | **321** | **151** |

## Directory Structure

Each category contains a `dataset.json` manifest and a `spreadsheet/` directory. Paths in the manifest are relative to the corresponding category directory.

### `dataset.json` as the Canonical Manifest

Each category has exactly one category-level `dataset.json`; it contains one record per task. This file is the authoritative source for task IDs, prompts, input workbooks, gold workbooks, and evaluation metadata. File and directory names are organized for readability, but users should always resolve a task through its manifest record rather than construct paths manually.

```text
nips2026_submit-v2/
|-- README.md
|-- Debugging/
|   |-- dataset.json
|   `-- spreadsheet/
|       |-- 01/
|       |   |-- 01_01_input.xlsx
|       |   |-- ...
|       |   `-- 01_golden.xlsx
|       `-- ...
|-- Financial_Model/
|   |-- dataset.json
|   `-- spreadsheet/
|       |-- 01/
|       |   |-- 01_01_input.xlsx
|       |   |-- ...
|       |   `-- 01_golden.xlsx
|       `-- ...
|-- Template/
|   |-- dataset.json
|   `-- spreadsheet/
|       |-- 01/
|       |   |-- 01_01_input.xlsx
|       |   `-- 01_01_golden.xlsx
|       `-- ...
`-- Visualization/
    |-- dataset.json
    `-- spreadsheet/
        |-- Task 96/
        `-- ...
```

Debugging, Financial_Model, and Template use the same numeric naming scheme. A numbered directory identifies a workbook group, and an input file uses the form `<group>_<task>_input.xlsx`. Shared references use `<group>_golden.xlsx`; Template references are task specific and use `<group>_<task>_golden.xlsx`. The numeric prefixes follow the record order in `dataset.json`.

Visualization is the only exception: its directories and files retain the human readable `Task <id>` naming scheme because each task has its own input workbook, reference workbook, and rubric set. The manifest is the authoritative index for all categories; users should resolve files through the manifest rather than infer task identity from filenames.

## Prompts and Manifest Fields

All 321 tasks include a complete, non-empty prompt in the `instruction` field. No category omits prompts, and prompts are not distributed as separate text files. The prompt, input path, gold path, and evaluation metadata for each task are stored together in its category-level `dataset.json` record. Common fields are:

| Field | Description |
|---|---|
| `id` | Task identifier within the category. |
| `instruction` | Complete prompt supplied to the spreadsheet agent. |
| `spreadsheet_path` | Relative path to the input workbook. |
| `golden_response_path` | Relative path to the completed or corrected reference workbook. |
| `answer_position` | Worksheet ranges used by the cell based evaluator. |

Debugging, Financial Model, and Template records include `answer_position`. Visualization records include `task_id`, `category`, and `criteria`; `criteria` contains the expert designed rubric used by the VLM evaluator.

## Input and Gold Mapping

| Category | Mapping rule |
|---|---|
| Debugging | Each numbered directory contains ten input variants named `<group>_<task>_input.xlsx` and one clean gold workbook named `<group>_golden.xlsx`. |
| Financial Model | Each numbered directory contains five incomplete task variants named `<group>_<task>_input.xlsx` that share one completed project workbook named `<group>_golden.xlsx`. |
| Template | Each task has its own input and gold pair, named `<group>_<task>_input.xlsx` and `<group>_<task>_golden.xlsx`. |
| Visualization | Each task has its own input workbook, gold workbook, prompt, and rubric set. The existing `Task <id>` directory and file names are retained. |

The exact input to gold mapping is always specified by `spreadsheet_path` and `golden_response_path` in the category manifest.

Each task record contains exactly one `golden_response_path`. In Debugging and Financial_Model, multiple task records intentionally point to the same gold workbook because they are controlled variants of one underlying workbook. In Template and Visualization, each task record points to its own gold workbook. Thus, there are no tasks with multiple gold paths; the apparent sharing occurs across task records and is explicitly represented by repeated `golden_response_path` values.

## Evaluation

- **Debugging, Financial Model, and Template:** generated workbooks are recalculated before cached values and the annotated worksheet ranges are compared with the reference workbook.
- **Visualization:** each criterion in `criteria` receives a binary VLM judgment, and the task score is the rubric pass rate.

## Loading the Dataset

```python
import json
from pathlib import Path

category_dir = Path("Financial_Model")
with (category_dir / "dataset.json").open("r", encoding="utf-8") as f:
    tasks = json.load(f)

for task in tasks:
    input_path = category_dir / task["spreadsheet_path"]
    reference_path = category_dir / task["golden_response_path"]
    assert input_path.exists()
    assert reference_path.exists()
```