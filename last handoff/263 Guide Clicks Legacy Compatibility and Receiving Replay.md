# Handoff 263 — Guide clicks, legacy compatibility and Receiving replay

## Goal and release outcome

Complete Plan 022 slice 4be-A under v4.11 D18 and D13. The reported button and Search behaviors now have passing packaged evidence on the November 22 candidate. Operator and final-candidate qualification remain open.

## Current verified state

Last verified 2026-11-23 Pacific. Code worker `invSys_edit_receipts`, branch `work/d18-picker-save-repro`, published through `2115f6b7`. Docs branch `work/perplexity-handoff218` started at `364a3e3` and publishes to `work/d18-picker-save-repro-docs`. No product VBA, XLAM or operational workbook changed. The five candidate hashes still match handoff 260. Excel is closed. Keep the existing untracked code reports, main checkout AGENTS edit and unrelated docs changes.

## Decisions and constraints

Justin authorized controlling, opening and closing Excel on this shared test computer. The workstation instructions require exiting on verified mouse/input failure. All five existing invSys Excel OPEN startup values were changed from the November 20 validation directory to `deploy/validation-search-extension-20261122` in the worker and read back correctly. Ordinary fresh-launch VBE path verification is still required before NAS review. The packaged guide fixture now checks each live project path before instrumenting its disposable copy.

## Evidence and traceability

| Gate | Result on the November 22 packages |
|---|---|
| Native picker/reader cycles | 92/92 in `slice4be-viewer-published-read/7c1e6ae386654fa3847e052fb96f7dff`. Both children opened and closed twice, one native click per action. Picker Create worked on one click. Matching caption returned the exact actions, no-match showed zero rows and clearing restored three. |
| Populated reader and recorded-run Create | 195/195 in `slice4be-viewer-published-read/f2d790cd393746a1884093c313cd3ee7`. One native Create click displayed the editor. Two guide versions filtered and restored. Reader Close and X disposed the window and instance. |
| Transfer, including legacy guides | 85/85 in `guide-transfer-controller/2a717e07b6c94523ae4ca2244febe380`. Both old schemas passed read, exact export, new linked revision, immutable predecessor, unassigned import and mandatory Subject on editing. Normal closure, settings and packages verified. |
| Early Receiving replay | 103/103 without captures in `receiving-replay-controller/4a9195dbc3f049c38af98f776bf4683a`. Settings restored, five files preserved, Excel closed and zero Excel crash/hang events through 45 seconds after completion. Controller keeps human B0 acceptance open. |
| Static surface | `guide-click-20261123-static` matched the preceding baseline at 392 components, 6,772 procedures, 146,601 lines, 1,265 candidates, 1,060 dynamic roots, 30 warnings and 9 literal/45 unresolved dynamic calls. |

These paths are relative to the worker's ignored `reports/runtime/`. A separate picker X trial at `slice4be-viewer-published-read/3bc6ff3e42ca4b61bf1f019b062dfb2a` returned 18 passes and one harness exception. Cursor restore succeeded according to Windows, while immediate readback differed on `btnChooseGuideActions`. The native X message was never sent. Cause is unknown. The reported 80010007 did not reproduce in either completed button suite.

## Do Not Repeat

The successful pointer preflight did not guarantee later restoration during real clicks. Avoid rerunning the complete mouse sequence merely to reach the separate X callback. Preserve the exact multiple-click and Search evidence.

## Assumptions to Re-verify

Confirm no Excel process before testing. Verify all five actual VBE paths in the fresh review host. Justin enters NAS credentials privately. Existing reviewer guide/recording counts may have changed.

## Open questions and blockers

The picker's X is unverified. The previous review used old running projects, so the reported error's behavior on the candidate NAS runtime remains unknown. Remaining gates include final launcher/Inventory/reusable Production, restart/binding, ordered live/full chain, and visible same-task Operations/Admin comparisons including incomplete runs. A detached chain checkout `invSys_4bea_20261123_chain` is prepared with the five candidate copies and prior qualified chain template. Its static chain wiring passed 13/13. The full chain has not run there.

## Immediate next action

On resumption, exercise native WM_CLOSE on the exact reopened picker using the existing packaged open/close handlers for fixture setup, then verify the window, loaded instance and next reopen before changing VBA.

## Critical references

D18-SEARCH-INPUT-02, D18-COURSE-01, D18-REPLAY-01, `Slice4beGuideActionCuration.ps1`, `Click-GuideControlNative`, `frmGuideActionPicker.UserForm_QueryClose`, `plan022_slice4be_remaining_acceptance.md` and `plan022_slice4be_operator_review.md`.
