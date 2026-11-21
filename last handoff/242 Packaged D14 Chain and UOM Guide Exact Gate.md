# Handoff 242: packaged D14 chain and UOM guide exact gate

Date: 2026-11-14 Pacific. Predecessor: handoff 241.

## Goal and release outcome

Complete invSys Release 1 Slice 4be-A under Architecture v4.11 D18-REPLAY-01 and Plan 022. The six D14 chain assertions, UOM guide identity gate and early B0 replay now have current-package technical coverage; human 4be-A acceptance remains open.

## Current verified state

Last verified 2026-11-14 Pacific: draft code PR16 (`work/d14-packaged-chain-coverage`) is pushed at `98602a6f4c780290e51380f3d6ad6bd3929ca997`, stacked on draft PR15 (`work/d18-recording-fixture-filter` at `0e988b2c1a8840334a9d930227e200b0cb2fdcac`). Both worktrees are clean. Main code checkout stays at `6d247543c557cca92fb6324f484b6ec2087e4527` with Justin's unrelated `AGENTS.md` edit preserved. Docs branch `work/perplexity-handoff218` carries Plan 022, Controls 1.756 and this handoff on draft PR5; unrelated dirty docs files remain untouched. Excel 16.0.20527.20196 is closed. The five validation XLAM hashes still match PR4-BASE-18.

## Decisions and constraints

No product VBA, XLAM, normative contract, ControlId or displayed wording changed. PR16 replaces only the obsolete source Create integration gate with current packaged Admin Create/Seed and Core snapshot/refresh assertions. Exact `System_Key` comparison is case-sensitive. Six static `ROW` warnings are reviewed only at exact source hashes; new paths or changed files still fail. Keep PR16 and PR5 draft until their remaining acceptance gates close. Do not combine separate runs into a claimed 4be-A GREEN.

## Evidence and traceability

- Focused packaged D14 v76 and exact-case v79 each passed 6/6 exact checks with natural Excel exit, restored settings and unchanged packages. The six cover managed headers/no `ROW`, seeded unique keys and `Condition=GOOD`, round trip, repeated Admin Seed without migration, custom column after refresh, and save/reopen. Ignored evidence: `reports/runtime/d14-packaged-focused-v79-20261114/`.
- Static warning classifier 5/5; chain contract 13/13. V77 chain passed 31/32 and exposed the static ratchet prerequisite. V78 passed 32/32 internally but Windows Application Error 1000 recorded an Excel access violation in `combase.dll` (`0xc0000005`, PID 24652); exact stage unproven. A fresh isolated v80 passed 32/32 with D14 gate, static ratchet, runtime five-package proof and local settings restoration; Excel exited naturally, package hashes matched, no modal or delayed Excel fault. See ignored `reports/runtime/release1-chain-d14-v80-20261114/qualification.json` and `slice14_results.v80.md`. This qualifies the technical chain regression on the current candidate, not the full acceptance matrix.
- Standalone v81 Inventory `Recording.Uom.Guide` passed 80/80 unique, exact ordered checks: the retained 74 names plus six `InventoryPartialChoices` names. Worker exit 0, natural Excel exit, restored settings, desktop error/lock/mouse counts 0, delayed Excel faults 0. It resolves v72 run 08's obsolete 74-name list. See ignored `reports/runtime/inventory-uom-guide-v81-20261114/run-08/qualification.json`.
- Read-only comparison found v20's B0 run used five XLAM bytes different from the current set. Standalone v82 on current packages passed 121/121 exact ordered Receiving replay checks, all four scope flags true and 10/10 proof-critical checks present. Natural Excel exit, settings and five package hashes preserved, desktop counters 0 and delayed faults 0. Representative setup, completed, stopped and resize captures were reviewed. The harness deliberately writes `B0Accepted=false`; human acceptance is separate. Ignored evidence: `reports/runtime/receiving-replay-current-v82-20261114/run-B0/qualification.json`.

## Do Not Repeat

Do not resume the old source Create harness by adding imports one at a time or replace it with packaged Create v62 alone; neither protects the six D14 identities. Do not call v78 an accepted chain, or call v72 run 08 product RED. The one-off v81 controller is ignored runtime evidence, not a changed product input or a new PR4 manifest.

## Assumptions to Re-verify

Recheck PR heads, five package hashes, Excel version/closure, current pointers and unrelated dirty files before another runtime or Git operation. The v78 Excel fault may recur; v80's clean run does not establish its cause.

## Open questions and blockers

Current-package early Receiving B0 technical replay is 121/121, with `B0Accepted=false` deliberately retained. Architecture v4.11 still requires visible operator comparison of How-To and Diagnostic for the same Operations and Admin tasks including an incomplete run, retained regressions, final build/compile/layout/header and saved-workbook binding evidence, and separate human acceptance. The intermittent v78 Excel fault remains unexplained.

## Immediate next action

Prepare the visible operator comparison for one Operations and one Admin task, including an incomplete run, against the current packaged candidate and record human acceptance separately.

## Critical references

Architecture v4.11 D13/D14/D18-REPLAY-01 (completion at lines 5226-5237); Plan 022; Controls 1.756; draft code PR16 and docs PR5; `tools/validate_release1_full_chain.ps1`; `tests/tooling/Run-D14PackagedCreate.ps1`; `tests/tooling/Invoke-D14PackagedGate.ps1`; PR4-BASE-18 manifest; ignored v79-v82 evidence.
