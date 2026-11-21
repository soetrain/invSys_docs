# Handoff 247: operator guide Save and editor latency

Date: 2026-11-16 Pacific. Predecessor: handoff 246.

## Goal and release outcome

Complete Slice 4be-A human and technical acceptance under Architecture v4.11 D18-REPLAY-01 and Plan 022 without regressing Release 1. A disposable Training guide now has a confirmed version 1, but editor responsiveness and the full How-To/Diagnostic comparison remain open.

## Current verified state

Last verified 2026-11-16: code worker `C:\Users\Justin\repos\invSys_edit_receipts` is `work/d18-picker-save-repro` at `8cad0ddc`, tracked-clean with the pre-existing untracked generated `tests/tooling/reports/`. Draft code PR17 targets `work/d14-packaged-chain-coverage`. Docs checkout is `work/perplexity-handoff218` at `5321644b` before this handoff commit; draft docs PR8 targets `work/excel-shutdown-timing`. The docs checkout retains Justin's unrelated deletion, edit and untracked files unchanged. Main code checkout `invSys_fork` remains at `6d247543` with Justin's unrelated `AGENTS.md` edit. Verify branch heads again before acting.

The final five-XLAM candidate is pinned to source `45d84ca8` in ignored `reports/runtime/4be-a-human-review-20261116b/addins/`; the same packages are in the worker's `reports/runtime/empty-holds-green2-20261116/addins/`. Excel PID 22272 was open and responsive on the review host, currently showing `Book1 - Excel`. Treat its open workbooks as potentially unsaved and leave them until the owner finishes or authorizes safe closure. The loaded XLAMs cannot be hashed through the ordinary file reader while Excel locks them. No new packaged test or build ran during this handoff's shift.

## Decisions and constraints

Architecture v4.11 is unchanged; no input-latency behavior change is approved. WHT7025AE remains Justin's disposable NAS Training warehouse. Its private PIN and test-only `ACTION_PATH_MAINT` grant do not qualify ordinary onboarding. The signed-out first-use proposal D8-TRAINING-ONBOARD-01 was cancelled. Preserve exact immutable `System_Key`, unknown columns, user workbooks and NAS source data. Controls is 1.764 and Plan 022 has the matching checkpoint.

## Evidence and traceability

- The operator clicked Save and saw **Published guide ... version 1**. Read-only NAS inspection found exactly one guide JSON version and no pending file. This proves one persisted guide Save, not guide reopening, comparison or 4be-A acceptance.
- Justin reports very slow action/step lists and title/instruction typing. `frmActionPathGuide` calls `modActionGuideDraft.WriteText` from each TextBox `Change`; `Guard` calls `modCuratedGuideSource.ReadSelected`, which rereads current tracking policy, `ACTION_PATH_MAINT` and loaded Events. `modActivityPolicy.ReadPolicy` opens and closes Config when it is not already open. This is a plausible cause, not a timed attribution. The test-only `Slice4beGuideDraft.ps1` and `Slice4beGuidePickerStoppedRun.ps1` changes at `3fadb025` exercise real packaged handlers for incremental typing, repeated action selection and step selection; both PowerShell files parse, but packaged RED is **unrun** because review Excel remains open. A local fixture may understate NAS latency; a passing local run alone would not dismiss the operator finding.
- Prior picker actual-handler RED 16 PASS / 1 FAIL at 10.39 s became rebuilt GREEN 17/17 at 0.84 s; direct curation retained 70/70. Shipping Holds final publication GREEN 86/86, five cold compiles and fresh Release 1 chain 33/33 with natural exit, preserved settings/package hashes and no Excel fault. Early Receiving B0 technical 121/121 was on the preceding picker build, with `B0Accepted=false`, not the final Shipping-reader build.
- Read-only ACE inspection of disposable `InventoryLog$` saw one event with a blank exact `System_Key`; exact ListObject scope and cause of the NAS Events warning remain unproven. No source data was changed. The prior `VBE7.DLL` crash remains unassigned.

## Do Not Repeat

Do not treat the guide version as proof of How-To/Diagnostic comparison, the test-only timing probe as behavioral RED, or the suspected guard cost as confirmed root cause. Do not run another Excel test/build while the review add-ins are open. Do not repair the blank-key event or reintroduce cancelled signed-out onboarding. Automatic approval review rejected recursive removal of generated `tests/tooling/reports/` as "blocked by policy"; leave it without rerouting.

## Assumptions to Re-verify

Confirm whether Justin is finished with the review host and whether `Book1` or any other workbook has unsaved work. After normal Excel closure, verify no Excel process remains and check package hashes, startup registration and the disposable NAS publication age. Recheck PR heads and Excel build before comparing historical timing.

## Open questions and blockers

Guide input latency needs focused packaged RED, a contract-preserving fix or an approved Architecture v4.11 amendment, then GREEN and visible operator confirmation. How-To/Diagnostic/Compare, complete/incomplete Operations and Admin runs, ordinary `ACTION_PATH_MAINT`/PIN setup, final-candidate B0 and retained packaged/layout/restart-binding gates remain open. Inventory Events publication needs exact-table read-only diagnosis or an approved disposable data strategy. 4be-A, 4be-B and Release 1 acceptance remain open.

## Immediate next action

After Justin's review Excel closes safely, run the focused `GuidePickerStoppedRunOnly` packaged timing test against the unchanged final five XLAMs and record exact named checks, per-field timing, fixture validity and natural Excel closure before proposing any product change.

## Critical references

Architecture v4.11 D18-REPLAY-01, 4be.5; Plan 022; Controls 1.764; code PR17; docs PR8; `tests/integration/plan022_slice4be_operator_review.md`; `tests/tooling/Slice4beGuideDraft.ps1`; `tests/tooling/Slice4beGuidePickerStoppedRun.ps1`; `tests/tooling/Test-Slice4beConfigCommands.ps1`; `frmActionPathGuide`; `modActionGuideDraft.Guard`; `modCuratedGuideSource.ReadSelected`; `modActivityPolicy.ReadPolicy`; ignored `reports/runtime/4be-a-human-review-20261116b/candidate.json`.
