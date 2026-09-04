# Evidence ledger

| Evidence | Immutable source | Local path | Role | License | Transformation | SHA256 | Supports | Risk |
|---|---|---|---|---|---|---|---|---|
| E1 | SpreadsheetBench-2 sample `Financial_Model/18_03` at `d8a14da1898e48d07c8a9ff8974d0effc9302ece` | `data/input_files/18_03_input.xlsx` | Solver-visible input | CC-BY-SA-4.0 | Original bytes | `fd095b2c0011d098aac379b6f8f928f79cc5a8ee4dcbdb81c6c5a88628ddf6c4` | Workbook state, labels, neighboring formula patterns | Upstream fixture is synthetic |
| E2 | Same dataset record | `adaptation/upstream-task.json` | Original request | CC-BY-SA-4.0 | Extracted record | n/a | Requested metrics and repair rules | Upstream wording audited below |
| E3 | Same dataset Golden | `adaptation/upstream-golden.xlsx` | Audit-only reference | CC-BY-SA-4.0 | Original bytes | `2406a01d43efa4fe61c12ec342d40fa12779a01ffd1b706297be7781c6bfd039` | Cross-check of independently justified formulas and structure | Not read by Judge |
| E4 | Dataset card and retained license | `adaptation/LICENSE-CC-BY-SA-4.0.txt` | Rights evidence | CC-BY-SA-4.0 | Retained text | `28a9529c7d0bb4dc51f4bf5c116a3d16ef247a052f7591466768ddf563fd1cf5` | Redistribution and ShareAlike basis | Human legal review remains required |

Every instruction requirement maps to E1+E2. Formula checks use visible labels, neighboring formulas and stated business rules; preservation and integrity checks use E1. E3 is used only to audit disagreements and the reference solution.
