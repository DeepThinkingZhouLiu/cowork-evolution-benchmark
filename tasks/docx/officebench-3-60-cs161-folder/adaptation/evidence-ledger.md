# Evidence ledger

| ID | Evidence personally read | SHA256 / finding | Adaptation use |
|---|---|---|---|
| E01 | OfficeBench commit | `b978b808667c32b52ce19a67ce1def1de9ae02b7` | Immutable upstream revision. |
| E02 | OfficeBench README | `04104d99107c2f109ea57d06f75541051efd4e7b5a8b2e41083315b3d21d5ed2` | Dataset context. |
| E03 | OfficeBench LICENSE | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4` | Apache-2.0 adaptation permission; byte-identical packaged copy. |
| E04 | OfficeBench evaluator implementation | `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90` | Established the original existence/keyword-only weakness. |
| E05 | `tasks/3-60/subtasks/0.json` / packaged source task | `4c648ffe87f47018910424bd20041b3a7eec829682c238d3215a517690eb64ae` | Preserves exact Query and evaluator. |
| E06 | Upstream/package workbook, all 124 × 9 cells | `f45c20ec06e621a6ad3b05549df9109f439c3bee09680d37c71d5087b5a418b3` | Sole record and assessment truth; both runtime copies match byte-for-byte. |
| E07 | Three prior DOCX reference copies | `70f2fb5ae362f7a71e56f95b17fe0d1fdac0e039035e93aa607fc3db8c8e70c7` each | Identical 123-paragraph dump; no native control structure. |
| E08 | Original instruction/rubric/Judge/solver/adaptation/metadata | full files read | Identified weak keyword scoring and missing row/score/structure validation. |
| E09 | Review policy, three frozen copies | byte-identical after validation | Defines block ranges, rules, sort, rounding, output count, and page contract. |
| E10 | Fresh input recomputation | 123 unique rows; 31/30/32/30 standings; 56 flagged | Judge truth basis. |
| E11 | Block recomputation | B1 `070f9935303d`; B2 `88abdb4e2a65`; B3 `546878816d0f`; B4 `14c19b854d00`; B5 `953a8453b231` | Closing control ledger truth. |
| E12 | Golden/fresh generation | SHA256 `9c96a13ef36c091861bf0d973028230c7090441c47adc7baeafe16fdb8da9471` | Byte-identical task-specific Oracle outputs. |
| E13 | Golden/fresh Judge repetitions | five 1.0 runs each; stable result JSON `ef1f3749776ad83f5e2a71edcbd6cc488d137618805fab0e7624df7887b73130` | Determinism and Oracle acceptance. |
| E14 | Directed negative suite | 28 packages; all non-Golden outputs below pass threshold; medium 0.78; two passes byte-identical | Failure-mode calibration. |
| E15 | Golden/fresh visual render | 11 pages each, all pages inspected | No blank, clipped, overlapping, unreadable, or wrongly oriented valid pages. |
| E16 | Nine representative negative renders | 9 packages / 10 changed pages inspected | Confirms score failures correspond to visible causal changes. |
| E17 | LibreOffice 7.3 round trip | DOCX `df0eeb02c3e27ac6fa40260b1ed53b998b87806776718fe483b88348611af9e5`; PDF `ba1b8503cb3be398344822eb56dc37b9ededda209ffa8bfdd19c8ba1102e6817`; Judge 0.83 | Business payload and layout survive; Word footer field encoding is not retained. |
| E18 | Harbor Docker Oracle A/B | trials `officebench-3-60-cs161-folder__Lri977a` / `officebench-3-60-cs161-folder__j6P8dq9`; rewards 1.0/1.0; 0/0 exceptions; checksum `cd77458e7fc2a39476b73b55f7ed99ee8f51f2123e633d7ad65dbfbb5381a26e`; digest `sha256:7fafdcfdcdd67beea063807c7008b9c1e74f872074e4b6c8cbd10b1658436ab9` | Same-snapshot container replay and artifact stability. |

The Judge derives truth from E06 and E09. E12 is audit evidence only and is never read by scoring code.
