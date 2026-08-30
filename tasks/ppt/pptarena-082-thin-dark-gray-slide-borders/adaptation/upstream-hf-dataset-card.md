---
license: mit
pretty_name: PPTArena
language:
- en
task_categories:
- other
tags:
- powerpoint
- pptx
- office
- presentations
- agentic-editing
- benchmark
- document-editing
- python-pptx
size_categories:
- n<1K
configs:
- config_name: default
  data_files:
  - split: test
    path: evaluation_pairs_refined.json
---

# PPTArena

**PPTArena** is a benchmark for **agentic PowerPoint (`.pptx`) editing**. Each case
gives a natural-language editing **instruction** and a **source deck**; a system must
edit the source so the result matches a human-authored **ground-truth deck**. It
measures whether agents can perform realistic, fine-grained slide edits — translation,
restyling, layout fixes, chart/table edits, transitions, master-level changes, and
more — while preserving everything they were *not* asked to change.

- **Paper:** [arXiv:2512.03042](https://arxiv.org/abs/2512.03042)
- **Code & web app:** https://github.com/michaelofengend/PPTArena
- **Leaderboard** https://ppt-arena.onrender.com/evaluation
- **License:** MIT
- **Cases:** 100 paired editing tasks (100 source + 100 ground-truth decks)
- **Modality:** binary Office files (`.pptx`)

## Dataset structure

```
PPTArena/
├── evaluation_pairs_refined.json   # 100 cases: prompt + (original, ground_truth) pairs
├── Original/        *.pptx         # 100 source decks to be edited
└── GroundTruth/     *.pptx         # 100 human-authored target decks
```

> [!NOTE]
> The `.pptx` files are binary Office documents. Use `huggingface_hub.snapshot_download`
> to fetch them (see **Usage**). `datasets.load_dataset` is **not** the intended access
> path for the decks — it would return raw bytes rather than usable files. The JSON
> index is what joins each prompt to its file pair.

## The index: `evaluation_pairs_refined.json`

A JSON array of 100 objects. Each object has:

| Key | Type | Description |
|---|---|---|
| `name` | string | Human-readable case id, e.g. `"Case 1: Translate to English But Keep French"`. |
| `prompt` | string | The editing instruction given to the agent. |
| `style_target` | string | Detailed description of the intended result (used for evaluation). |
| `original` | string | Path to the source deck, e.g. `"Original/Foo_TestA.pptx"`. |
| `ground_truth` | string | Path to the target deck, e.g. `"GroundTruth/Foo_GroundTruthA.pptx"`. |
| `category` | list[string] | High-level category, e.g. `["Content"]`, `["Layout"]`. May contain more than one. |
| `edit_type` | string | Fine-grained edit type, e.g. `"Text & Typography"`, `"Charts"`. |
| `enhancement_notes` | string | Notes on how the case/target was refined. |

### Categories (cases per category; a case may have more than one)

| Category | Count |
|---|---|
| Content | 67 |
| Layout | 29 |
| Styling | 29 |
| Structure | 15 |
| Interactivity | 4 |

### Edit types (cases per type)

Text & Typography (29), Charts (10), Images & Pictures (10), Theme & Background (9),
Alignment/Distribution/Z-order (8), Slide/Section Management & Footers (8), Tables (8),
Shapes & Drawing (4), SmartArt & Diagrams (4), Slide Layout & Placeholders (3),
Accessibility & Semantics (2), and one case each of Slide Transitions, Hyperlinks &
Action Settings, Template & Master-Level Edits, Audio & Video, and Object Animations.

## Usage

```python
# pip install -U huggingface_hub
from huggingface_hub import snapshot_download
import json, os

# Download the whole benchmark (decks + index) into the HF cache.
local = snapshot_download(repo_id="<username>/PPTArena", repo_type="dataset")

# Load the 100-case index.
cases = json.load(open(os.path.join(local, "evaluation_pairs_refined.json")))

# Resolve a single case to absolute .pptx paths.
case = cases[0]
original     = os.path.join(local, case["original"])      # the .pptx to edit
ground_truth = os.path.join(local, case["ground_truth"])  # the target .pptx
print(case["prompt"])
# `original` and `ground_truth` are real .pptx files: open them with python-pptx,
# feed them to an agent, etc.

# Optional: fetch only the index + source decks (skip ground truth) to save bandwidth:
# snapshot_download(repo_id="<username>/PPTArena", repo_type="dataset",
#                   allow_patterns=["*.json", "Original/*"])
```

## Evaluation protocol

For each case: apply `prompt` to the deck at `original`, producing a candidate deck,
then compare the candidate against the deck at `ground_truth` (the `style_target` field
describes the intended outcome in detail). The benchmark rewards making exactly the
requested edit while preserving unrelated content, layout, and formatting. The paper
uses a dual VLM-as-judge pipeline scoring instruction-following and visual quality.

## Citation

```bibtex
@article{ofengenden2025pptarena,
  title   = {PPTArena: A Benchmark for Agentic PowerPoint Editing},
  author  = {Ofengenden, Michael and Man, Yunze and Pang, Ziqi and Wang, Yu-Xiong},
  journal = {arXiv preprint arXiv:2512.03042},
  year    = {2025}
}
```

## License

Released under the **MIT License**.
