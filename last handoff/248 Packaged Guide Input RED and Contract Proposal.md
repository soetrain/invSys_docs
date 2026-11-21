# Handoff 248: packaged guide input RED and contract proposal

Date: 2026-11-16 Pacific. Predecessor: handoff 247.

## Goal and release outcome

Complete Slice 4be-A under Architecture v4.11 D18-REPLAY-01 and Plan 022 while preserving Release 1. One disposable Training guide version 1 is published; a qualified guide-input RED now blocks operator acceptance.

## Current verified state

Last verified 2026-11-16: code worker `C:\Users\Justin\repos\invSys_edit_receipts` is `work/d18-picker-save-repro` at `f0f97bab`, tracked-clean except pre-existing generated untracked `tests/tooling/reports/`; draft code PR17 targets `work/d14-packaged-chain-coverage`. Docs checkout is `work/perplexity-handoff218` at `ce4d683c` before this handoff/proposal commit; draft docs PR8 targets `work/excel-shutdown-timing`. Preserve its unrelated deletion, edit and untracked files. Main code checkout `invSys_fork` remains at `6d247543` with Justin's unrelated `AGENTS.md` edit.

Excel PID 22272 closed naturally after a read-only check that the review host and `Book1` were both marked Saved. The focused test's Excel also exited naturally; no Excel remained after the run. The five unchanged XLAMs in worker `reports/runtime/empty-holds-green2-20261116/addins/` matched the final review `candidate.json` hashes before and after dispatch; the Application audit found no Excel Error 1000/Hang 1002 in the test window. Recheck Excel before another test/build.

## Decisions and constraints

Architecture v4.11 contains **proposed, unapproved** `D18-GUIDE-INPUT-01`; Plan 022 and Controls 1.765 mark the same proposed state. Do not implement buffering until Justin approves the normative amendment. The proposal treats keystrokes as local unsaved form input and requires guarded Core flush before dependent actions and Save; Save retains its own full guard. Review refinements require dirty step text to flush against its previous editing StepId before switching/reordering/removing/reloading, and Cancel/X to discard before any focus-exit flush, including stale authority. The cancelled signed-out onboarding proposal stays cancelled. Preserve captured session/warehouse, exact `System_Key`, unknown columns, source bytes and unrelated user work.

## Evidence and traceability

- Qualified actual-handler RED on source `45d84ca8` packages: **24 PASS / 3 FAIL**, with five instrumented compile passes, valid disposable recording/Events fixtures, exact authored text preserved, both immutable guide Save routes, Cancel and source-byte preservation passing. Failed checks: `GuidePickerStoppedRun.txtGuideName.IncrementalInputResponsive`, `.txtGuideInstructions.IncrementalInputResponsive`, `.txtGuideStepInstruction.IncrementalInputResponsive`. The 20/25/22-character fields took 19.89/26.36/23.95 seconds against 2.5-second bounds. First/repeated picker toggles and step selection were about one second and passed local bounds; Justin still reports slow NAS lists. Excel exited naturally. Ignored raw evidence: worker `reports/runtime/slice4be-viewer-published-read/cf6b6fe9085d4da18cf6fe63f12c87a3/` (`red.json`, `guide-editor-input-timing.json`).
- Symptom -> likely path -> rule: per-character `frmActionPathGuide.StageText` calls `modActionGuideDraft.WriteText` -> `Guard` -> current policy/Auth/loaded Events; `modActivityPolicy.ReadPolicy` opens/closes Config if not already open. This is source-supported suspicion, not timed attribution within individual guards. Governing rule: Architecture v4.11 4be.5 captured binding and guards; D13 actual-handler RED/GREEN.
- Earlier focused picker GREEN 17/17, direct curation 70/70, final Shipping Holds publication GREEN 86/86 and full Release 1 chain 33/33 remain separate. Early Receiving B0 121/121 was on the preceding build with `B0Accepted=false`, not final-candidate acceptance. The operator's saved guide version 1 is confirmed; How-To/Diagnostic/Compare and guide reopen are untried.

## Do Not Repeat

Do not count text-preservation passes as responsiveness GREEN, a local one-second selection as NAS usability acceptance, or the proposed buffered-entry contract as approved. Do not weaken policy/capability checks silently or repair the disposable blank-key Inventory event. Do not reroute the policy-blocked generated-report deletion.

## Assumptions to Re-verify

Confirm the proposal's step-switch/focus and stale-context behavior against current form handlers before implementation. Recheck package hashes, Excel build, NAS performance, publication age, PR heads and worktree status before the next runtime gate.

## Open questions and blockers

Justin's approval of `D18-GUIDE-INPUT-01` is required for buffering implementation. The list-selection complaint still needs visible NAS evidence after any input fix. Ordinary `ACTION_PATH_MAINT`/PIN setup, final-build B0, retained packaged/layout/restart-binding gates, complete/incomplete Operations and Admin comparison, 4be-A, 4be-B and Release 1 acceptance remain open.

## Immediate next action

After approval, extend the focused actual-handler test for dirty step switch/reorder/removal, tags, blur, expectation entry/return, refresh/rebind, Save while focused with exact persisted text/StepId, and dirty Cancel/X with stale authority and zero `WriteText`, then implement only the approved buffering contract and run packaged GREEN.

## Critical references

Architecture v4.11 `D18-GUIDE-INPUT-01` and 4be.5; Plan 022; Controls 1.765; code PR17; docs PR8; `tests/tooling/Slice4beGuidePickerStoppedRun.ps1`; `tests/tooling/Slice4beGuideDraft.ps1`; `tests/integration/plan022_slice4be_operator_review.md`; `frmActionPathGuide.StageText`; `modActionGuideDraft.Guard`; `modCuratedGuideSource.ReadSelected`; `modActivityPolicy.ReadPolicy`; `reports/runtime/4be-a-human-review-20261116b/candidate.json`.
