# Handoff 262 — Native guide click reproduction interrupted

## Goal and release outcome

Qualify first-click guide creation, repeatable child closure and picker Search under Plan 022 slice 4be-A, D18-SEARCH-INPUT-02 and D13.

## Current verified state

Last verified 2026-11-23 Pacific. Code worker `invSys_edit_receipts` branch `work/d18-picker-save-repro` at `b5f210b2` adds native regression checks. Docs branch `work/perplexity-handoff218` began at `52c7d40`. All five November 22 review package hashes matched handoff 260 before and after the test. Zero Excel processes remained. Code has the pre-existing untracked tooling reports, the main checkout has its AGENTS edit, and unrelated docs changes listed by Git remain local.

## Decisions and constraints

The new observations fit the existing 4be-A form contract. D13 requires behavioral RED before changing VBA. The workstation instruction requires exiting on verified mouse/input failure. The startup registry still points all five invSys entries to the November 20 validation set.

## Evidence and traceability

The operator reported double-click entry, a second-open Close failure on both children, unresponsive picker Search and RPC error 80010007 at `frmGuideActionPicker.mCancel_Click` calling `mOwner.CloseGuideActionPicker`. The owning form's lifetime at the time of the error is unknown.

On isolated saved copies of the November 22 packages, compilation passed for all five projects. The new native test delivered one click to open an empty Published guides reader and one to close it, then verified both the native window and loaded MSForms instance were gone. On the second opening the cursor helper reported a successful Windows restore call but read back a different position. The test exited with 16 passes and one harness exception before checking second-open closure. Root cause of the cursor discrepancy is unknown. Local result in the worker is `reports/runtime/slice4be-viewer-published-read/2f22390389d34adeb315cfb664db9284/red.json`. No product RED or full GREEN was obtained. The test did not reach picker Search or either Create guide assertion.

## Do Not Repeat

Keep the November 20 operator review, November 22 small-fixture gates and this interrupted native run distinct. Add-ins registration alone cannot identify the executing VBA projects.

## Assumptions to Re-verify

Verify stable native mouse input and Excel closure before the next packaged run. For a visible NAS review, verify all five live VBE project paths in a fresh process and have Justin sign in privately.

## Open questions and blockers

It remains unknown whether either child fails on its second Close or Create needs two clicks on the November 22 candidate. Published guides Search in the operator's empty library is inconclusive. The package, larger-list NAS and overall 4be-A gates remain open.

## Immediate next action

With stable input, rerun `Test-Slice4beConfigCommands.ps1` on `deploy/validation-search-extension-20261122` with `-CheckViewerPublishedRead -GuideDraftOnly -CompileEvaluationProbesForTest -ViewerStartupPackageStateForTest SavedCopies -CheckGuideSearchNative -VisibleExcelForTest -GuideActionCurationOnly -Phase RED`, then run the populated reader/recording Create fixture with `-CheckGuideExpectation` in place of `-GuideActionCurationOnly`.

## Critical references

Handoffs 260–261, `Slice4beGuideActionCuration.ps1`, `Slice4beGuideDraft.ps1`, `Slice4beGuideInputNative.ps1`, `frmActionPaths.CloseGuideActionPicker` and `frmActionPathLibrary.mClose_Click`.
