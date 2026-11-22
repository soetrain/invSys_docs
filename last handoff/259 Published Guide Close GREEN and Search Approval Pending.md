# Handoff 259: Published guides Close GREEN; search approval pending

Date: 2026-11-21 Pacific. Predecessor: handoff 258.

## Goal and release outcome

Complete Slice 4be-A under Architecture v4.11 and Plan 022. This increment restores the existing D18 Published guides Close contract; both slow Search controls and final operator acceptance remain open.

## Current verified state

Last verified 2026-11-21: code worker `C:\Users\Justin\repos\invSys_edit_receipts` branch `work/d18-picker-save-repro` at `35cc781d`, draft [PR17](https://github.com/soetrain/invSys_fork/pull/17), base `work/d14-packaged-chain-coverage`. Docs branch `work/perplexity-handoff218` at `7c539593`, published to draft [PR8](https://github.com/soetrain/invSys_docs/pull/8), base `work/excel-shutdown-timing`. Excel is closed. Code still has pre-existing untracked `tests/tooling/reports/`; docs still have unrelated guidance/handoff edits and untracked files. Preserve them and the unrelated `invSys_fork/main` AGENTS.md edit.

## Decisions and constraints

The owner designated WHT7025AE as disposable NAS training. No PIN is in the repos or available to this handoff. D18-SEARCH-INPUT-02 is **proposed, not approved** in Architecture v4.11: local per-keystroke filtering for Choose tracked actions and Published guides would explicitly replace D18-COURSE-01's per-key Core check for Published guides. Do not implement that behavior until approval. D18-RECORDING-LIST-01 is separately pending. Published guides Close uses the existing reader contract, with no new architectural rule.

## Evidence and traceability

- Native-X baseline: 176 PASS / two Close FAIL, `reports/runtime/slice4be-viewer-published-read/d932ddfaf8794faa9b3c5221dc5833ca/red.json`.
- Deferred owner cleanup after native QueryClose: five packaged XLAM compiles and 178/178 actual-handler GREEN, including Close button, native X, loaded-reader disposal, Viewer closure, policy, layout, exact guide/source bytes and expectation checks. Excel exited naturally; no matching Application Error 1000/1002 in that run. Evidence: `reports/runtime/slice4be-viewer-published-read/4865df0c8e02428ca85f1ea150b704c0/green.json`.
- Static maintenance: `reports/runtime/action-reader-close-static-roots-20261121/` reports 392 components, 6,770 procedures, 1,265 candidates, 27 oversized-module ratchets, and 30 warnings. The one added registered dynamic root is the Excel `OnTime` callback; no new warning. Plan 022 and Controls 1.795 record the Close checkpoint. Runtime reports are ignored and not committed.
- One earlier same-build run passed native X and Viewer closure but later stopped at an Excel-busy COM error in expectation reading; the unchanged repeat passed 178/178. This does not qualify broader 4be-A or Release 1.

## Do Not Repeat

Synchronous owner unload inside the native `QueryClose` leaves the reader open. Releasing layout references inside that event caused a VBE7.DLL crash in two discarded trials; both source attempts were reverted. The tested fix schedules owner cleanup after the event returns.

## Assumptions to Re-verify

Recheck Excel's closed state, package build and PR heads before more runtime work. The user's signed-in disposable review session may no longer be open. Visible NAS confirmation of Close remains outstanding.

## Open questions and blockers

Explicit owner approval of D18-SEARCH-INPUT-02 is pending. Both Search actions in Choose tracked actions and Search Subject/Title/Tags/ID in Published guides remain slow under the current contract. Slice 4be-A, 4be-B and Release 1 acceptance remain open.

## Immediate next action

After D18-SEARCH-INPUT-02 approval, add and run focused actual-handler RED tests for both searches against the current packaged XLAMs before changing their implementation.

## Critical references

Architecture v4.11 D18-COURSE-01 and proposed D18-SEARCH-INPUT-02; Plan 022; Controls 1.795; `frmGuideActionPicker.mSearch_Change`, `frmActionPathLibrary.mSearch_Change`/`UserForm_QueryClose`, `frmActionPaths.ClosePublishedGuidesIfCurrent`, `modGuideReaderClose.Flush`, `Slice4beGuideLibrary.ps1` and `Slice4beGuideDraft.ps1`.
