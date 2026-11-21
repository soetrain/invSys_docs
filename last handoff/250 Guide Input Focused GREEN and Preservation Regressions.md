# Handoff 250: guide input focused GREEN and preservation regressions

Date: 2026-11-19 Pacific. Predecessor: handoff 249.

## Goal and release outcome

Complete Slice 4be-A under Architecture v4.11 and Plan 022 without Release 1 regressions. Approved `D18-GUIDE-INPUT-01` now has focused packaged handler GREEN, but its broader guide regression and 4be-A acceptance remain open.

## Current verified state

Last verified 2026-11-19: code worker `C:\Users\Justin\repos\invSys_edit_receipts`, branch `work/d18-picker-save-repro`, commit `4cfdd1c0` on draft code PR17. It contains the guide-input candidate and test tooling. Only pre-existing `tests/tooling/reports/` remains untracked. Docs branch `work/perplexity-handoff218` recorded the Plan 022 and Controls 1.768 checkpoint at `d0a7fb4d` before this handoff. Draft docs PR8 remains open. Preserve its unrelated dirty and untracked files. Main code checkout remains at `6d247543` with Justin's `AGENTS.md` edit. Excel count was zero after the retained test and after the later unexpected reboot; recheck before any test or build.

## Decisions and constraints

Justin approved `D18-GUIDE-INPUT-01` on 2026-11-17. The candidate buffers Name, Tags, Instructions and Step Instruction locally; keeps the previous editing StepId; flushes through typed Core at dependent actions, Save, Tab and cross-field mouse focus; and discards dirty Cancel/X without a Core write. Runtime-created MSForms controls did not deliver usable `Exit` or `AfterUpdate` callbacks in the focused packaged experiment, so Tab `KeyDown` and cross-field `MouseDown` supplement explicit action boundaries. This is an implementation of the approved rule, not a new contract. Do not deploy to `deploy/current` or call 4be-A accepted.

## Evidence and traceability

- Unchanged-package actual-handler RED is in handoff 249. The subsequent qualified mouse-focus RED was **41 PASS / 2 FAIL**: Name did not flush on mouse transition and the later step switch wrote two pending fields. Ignored runtime report is under `reports/runtime/slice4be-viewer-published-read/` in the worker.
- Isolated five-XLAM candidate `reports/runtime/guide-input-green6-20261117/addins/` passed five instrumented compiles and focused actual-handler **43/43 GREEN**. `invSys.Operations.xlam` SHA256 `6FAF34132B35B3A10FBA95F9A6D01B7AB20E2DF572C3BE22ED47740679C4AFE3`; Core SHA256 `43DF458CC4A6533839567ACC810F552A96ABD3CCCC815D2DD783D415C7BDBC01`. Report `reports/runtime/slice4be-viewer-published-read/72573d40798d42e5821a5a1c5e5d26e5/` records one write each for Name, Tags, Instructions blur and the dirty step switch, exact persisted text/StepId, both Save routes, and zero-write stale dirty Cancel/X.
- A separate retained saved-copy Viewer/guide run on the same candidate passed **169/171**. Report `reports/runtime/slice4be-viewer-published-read/5bc84b904b1447e9aa6b05a9f0b8e55b/green.json` names only `GuideLibrary.ReadsPreserveGuideSourceAndActivityBytes` and `GuideExpectation.AllActionsPreserveRecordedExpectationsAndActivity` as failed. These compound checks compare source, guide and activity pins; the changed predicate/file is not yet known. Viewer layout, Guide Draft reorder/removal/layout, Save, Library and most Expectation checks passed. This is not broad GREEN.
- Windows logged an unclean shutdown after the 2026-11-17 run and booted again on 2026-11-19. The checked 2026-11-17 11:00–12:00 Pacific Application Error 1000/1002 window had no Excel-named fault. Cause and relationship to the test are unproven.

## Do Not Repeat

Do not treat `Exit`/`AfterUpdate` on the runtime-created textboxes or programmatic `SetFocus` plus assigned `Value` as a qualified ordinary-blur test. Native click plus keystrokes and Tab gave the focused evidence. A cursor mismatch immediately after `SendInput` was a test cleanup race; the helper now waits for mouse-up before restoring the pointer. Do not combine the 43/43 focused result with the separate 169/171 run into an acceptance total.

## Assumptions to Re-verify

Recheck Excel closure, active RDP input/cursor movement and restoration, candidate package hashes, PR heads, repo status and whether the two preservation failures reproduce. The Windows reboot could have changed desktop or Excel state.

## Open questions and blockers

The exact source/guide/activity pin behind the two retained failures is unresolved. Dirty reorder/removal, expectation entry/return, refresh/rebind and stale session/policy/capability refusal need focused input-order coverage beyond existing retained checks. Static maintenance, live-role, full Release 1 chain, restart, visible NAS usability and operator How-To/Diagnostic/Compare acceptance remain open. Slice 4be-A, 4be-B and Release 1 remain incomplete.

## Immediate next action

Add a sanitized per-predicate pin diagnostic to the disposable Guide Library and Guide Expectation fixtures, then rerun their packaged regression on the unchanged candidate to isolate the two preservation failures before changing product behavior.

## Critical references

Architecture v4.11 `D18-GUIDE-INPUT-01`; Plan 022; Controls 1.768; code PR17; docs PR8; `frmActionPathGuide.frm`; `modGuideEditor.bas`; `Slice4beGuidePickerStoppedRun.ps1`; `Slice4beGuideInputNative.ps1`; `Slice4beGuideLibrary.ps1`; `Slice4beGuideExpectation.ps1`; the two ignored runtime reports above.
