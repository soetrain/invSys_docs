# Handoff 264 — Picker native close and Release 1 chain

## Goal and release outcome

Complete 4be-A under D13 and D18. The reopened picker now closes with the native X, and the five-package build passed the ordered Release 1 chain. Ordinary Excel startup and the visible review remain open.

## Current verified state

Last verified 2026-11-23 Pacific. The code worker `invSys_edit_receipts` on `work/d18-picker-save-repro` is published through `4a809b81`. Docs use `work/perplexity-handoff218`, published to `work/d18-picker-save-repro-docs`. The previous docs commit was `042622b`. The current five-file candidate is `deploy/validation-guide-picker-close-20261123`. The review Excel is open with an informational modal. Confirm closure before another Excel test or build. Preserve unrelated changes in both primary checkouts. Prior PRs remain closed.

## Decisions and constraints

Justin authorized Excel control and delegated the technical assessment. The workstation's five existing OPEN entries now point to the November 23 candidate. No shared deployment changed. D18 still requires visible NAS review. Earlier live review used different running VBA projects. Verify actual project file paths before signing in. The existing NAS test account uses private on-device sign-in.

## Evidence and traceability

| Check | Result |
|---|---|
| Picker X after reopen | Previous candidate 14 passes and two failures, including the downstream fixture stop. `UserForm_QueryClose` cancelled X and requested a nested unload through its owner. The corrected picker finishes its own close, releases selection and clears the parent's exact-instance reference. Rebuilt candidate 77/77 with native X, subsequent reuse, Cancel, parent and Viewer closure, and source integrity. |
| Cold startup and explicit compilation | All five XLAMs compiled in the packaged compile test. |
| Warehouse workflows | Release 1 ordered chain 33/33, including restart reconciliation. Packaged live role 48/48. Settings restored, five hashes preserved, zero Excel crash/hang events. Static chain wiring 13/13 and control surface 6/6. |
| Maintenance | 392 components, 6,773 procedures, 146,601 lines, 1,265 candidates, 1,060 dynamic roots, 30 warnings, 9 literal and 45 unresolved dynamic calls. One new typed owner notification accounts for the added procedure. |
| Ordinary desktop launch | Opening the disposable review host read-only surfaced Excel's “This workbook is currently referenced by another workbook and cannot be closed.” Captured OK via UI Automation. The launch failed before sign-in and did not verify live VBA paths. Four package projects point to the newly built Core, while the first two registry entries name Admin then Core. Startup order is a hypothesis. The notice has no matching literal in product or test sources. |

Worker-relative evidence: `reports/runtime/slice4be-viewer-published-read/1635a830e66748a3a026919cca888995/red.json`, `reports/runtime/slice4be-viewer-published-read/795c18e39a184ecb859fb7192a1af976/green.json` and `reports/runtime/4bea-picker-finish-20261123/{package-pins.json,chain-summary.json,ordinary/modal-20261123T195811Z.json}`. The detached checkout `invSys_4bea_20261123_chain` retains the generated integration results. Share only sanitized results.

## Do Not Repeat

Count the ordinary launch as failed regardless of any later acknowledgement. Avoid starting another Excel process while the notice remains. Do not infer project paths from registry entries or use the November 22 small-list native timings as November 23 large-list qualification.

## Assumptions to Re-verify

Recheck the open notice, desktop and five package hashes. The review host is disposable; the November 23 attempt opened it with Excel's read-only switch. Full startup ordering and the NAS user experience remain unverified.

## Open questions and blockers

The workstation instructions require owner approval to acknowledge the unexpected notice. The operator has been asked. Remaining candidate gates include larger-list native Search, repeated native button cycles, saved-workbook binding, final launcher/Inventory/reusable Production/layout, guide transfer, early Receiving B0 and visible complete/incomplete Admin and Operations comparisons. Record human acceptance separately.

## Immediate next action

After the captured Excel notice is cleared under owner approval and the test instance is closed, verify a fresh read-only launch with Core ordered first and inspect all five actual project paths before continuing the larger-list Search test.

## Critical references

`frmGuideActionPicker.UserForm_QueryClose`, `frmActionPaths.ForgetClosedGuideActionPicker`, `Slice4beGuideActionCuration.ps1`, `-GuidePickerScaleOnly`, `plan022_slice4be_remaining_acceptance.md`, `plan022_slice4be_operator_review.md`, D18-SEARCH-INPUT-02 and D18-REPLAY-01.
