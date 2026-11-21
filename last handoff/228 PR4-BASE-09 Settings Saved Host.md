# Handoff 228: PR4-BASE-09 Settings saved host

Date: 2026-11-12 Pacific. Predecessor: handoff 227.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. The current Settings gate is unqualified; a saved-workbook test host now matches the established R1 operator model, but full packaged GREEN and human acceptance remain open.

## Current verified state

Last verified 2026-11-12: code `main` is `6d247543` with Justin's unrelated `AGENTS.md` edit preserved. Execution worker `C:\Users\Justin\repos\invSys_edit_receipts` was restored clean to `work/d8-inv-edit-receipts` at `550d9ff0`; draft code PR14, `work/d18-settings-saved-host` at `d94768c7`, is stacked on draft PR13 and remains in its separate clean worktree. Docs `work/perplexity-handoff218` has Plan 022 and Controls **1.740** at `363c7f2`; this handoff follows. Existing unrelated docs deletion, edit and untracked files remain untouched. Excel is closed. Recheck all state before runtime.

## Decisions and constraints

- Architecture v4.11 already requires a saved `.xlsm`/`.xlsb` operator workbook reopened under the common XLAM session and retains add-in-only resource stress as an unresolved diagnostic. No architecture, Plan contract, product VBA, XLAM, ControlId or wording change was made.
- v32 and separately approved v32.1 are attended harness failures, not product RED or Settings GREEN. Both authorizations are consumed. The prepared PR4-BASE-09/v33 run requires a separate one-use approval.
- Keep natural Excel closure, exact pins, desktop monitor, input preservation and the user's 2026-11-12 **21:08 UTC** shift stop. Stop on lockscreen/verified input failure or required approval. Do not rerun prior diagnostics.

## Evidence and traceability

- v32/v32.1 each printed 134 PASS, one fixture-target exception (`WH_CONFIG_INVALID`) and left 68 of 202 Settings checks unrun after captured Excel insufficient-memory dialogs. v32.1 reached 10,000 GDI objects with about 40 GiB free virtual memory; its delayed audit found nine `GDIObjectLeak` reports naming Excel. Both exited naturally and preserved settings/pins. `reports/runtime/settings-policy-v32-1-pr4-base-08-20261112` contains closure/modal evidence.
- Test-only packaged probe `reports/runtime/4be-config-resource-probe-20261112`: with **zero saved workbooks**, 80 successful `modConfig.LoadConfig` calls raised GDI 1,036→3,606 and `XLMAIN` 20→100; Settings form cycles hit the 5,000-object safety cap. `Workbooks.Count` returned to zero. This reproduces the existing add-in-only host condition.
- Distinct test-only probe `reports/runtime/4be-config-saved-host-probe-20261112`: with one disposable saved `.xlsb` reopened under five packages, 80 Core loads held GDI 927–929 and `XLMAIN` 16; 20 Settings open/show/close cycles ended at 935 GDI with zero loaded forms. Worker returned 0, Excel exited naturally, and no new Excel Application Error 1000/1001/1002 was observed. This is focused diagnostic GREEN, not the 202-check gate. Plain `.xlsx`/`.xlsb` 40-cycle controls also stayed near baseline.
- Draft code PR14 `d94768c7` adds one guarded saved-host switch to `tests/tooling/Test-Slice4beConfigCommands.ps1`; PowerShell parse and diff check pass. PR4-BASE-09 manifest `6F867D13...` has 539 pins, differing from BASE-08 only in that test worker. The five XLAMs are unchanged. The v33 launcher `6702737E...`, expected list `0070059C...` and `codex-settings-policy-v33.md` are in `reports/runtime/settings-policy-v33-pr4-base-09-20261112`. Checks-only passed 539 pins/202 names with no Excel. No `session-authorization.md` or dispatch exists.
- Earlier standalone Inventory/Recording 4be-A results and technical B0 replay remain as in handoff 227 and Plan 022. `B0Accepted=false`; do not sum independent gates into acceptance.

## Do Not Repeat

Do not attribute the add-in-only GDI growth to the operator path, retry v32/v32.1, treat a fixture exception as product RED, raise Windows GDI quotas, or promote PR14 before packaged qualification. The v33 controller has a fixed **19:50 UTC start cutoff** and may not be edited in place to extend it.

## Assumptions to Re-verify

Recheck remote PR heads, clean worker state, five XLAM hashes, 539 pins, Excel closure, desktop/input, Excel build and available memory before dispatch. The saved-host probe was a focused comparison, not full Settings behavior; whether all 202 checks pass remains unknown.

## Open questions and blockers

Fresh one-run v33 approval and Settings packaged GREEN; remaining Event Viewer/detail, guide transfer and other Admin/Operations families; retained-formatting callers; human B0; live-role, restart/binding, full-chain and visible operator acceptance. Residual Execute slowness and the Microsoft-reported Excel exit defect remain separate.

## Immediate next action

Obtain approval for PR4-BASE-09 and one v33 run, detach the clean execution worker at `d94768c7`, rerun checks-only, then dispatch once before 19:50 UTC only if all gates pass.

## Critical references

Architecture v4.11 operator deployment model and D18-REPLAY-01; Plan 022 PR4-BASE-08 checkpoint; Controls 1.740; `tests/tooling/Test-Slice4beConfigCommands.ps1`; draft code PR14; docs PR5; local probe, manifest and v33 controller folders above.
