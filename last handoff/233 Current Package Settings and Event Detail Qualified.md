# Handoff 233: current-package Settings and Event Detail qualified

Date: 2026-11-12 Pacific. Predecessor: handoff 232.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. The current five-XLAM candidate now has standalone Settings and Event Detail GREEN; full Viewer, Action Path and Release 1 acceptance remain open.

## Current verified state

Last verified 2026-11-12: code `main` is `6d247543` with Justin's unrelated `AGENTS.md` edit untouched. Draft code PR14 is clean at `b8f94241`, stacked on PR13. The execution worker `C:\Users\Justin\repos\invSys_edit_receipts` is restored to clean `work/d8-inv-edit-receipts` at `550d9ff0`. Docs PR5 branch `work/perplexity-handoff218` has Plan 022 and Controls **1.745**; unrelated docs changes remain untouched. Excel is closed. No run is active. The owner's shift ends **21:08 UTC**.

## Decisions and constraints

- Architecture v4.11's saved operator workbook contract, five validation XLAMs, product VBA, ControlIds and displayed wording are unchanged. PR14 changes only test tooling.
- v35 expired before dispatch. v36 and v39 one-run approvals are consumed. v37 and v38 are consumed diagnostics, never to be retried. Future Excel runs need their own bounded approvals and frozen inputs.
- The `soetrain` docs branch `work/perplexity-handoff218` is an older PR6 tree. Docs PR5's current head is the `justinwj` fork branch; do not merge the older branch into current docs work.

## Evidence and traceability

- v36 PR4-BASE-10 Settings/policy/detail/preference: **202/202 exact PASS**, worker exit 0, controller `StopReason=null`, pinned restart proof and both Excel native exits verified, no extra Excel, settings restored, 539 pins preserved, no process termination/modal, delayed audit zero faults. The first-Excel stdout-quiet metric is inapplicable because the first process exits before later checks. Evidence: `reports/runtime/settings-policy-v36-pr4-base-10-20261112`.
- v37 stopped before Excel because the controller supplied a compile switch unsupported by the Event Detail selector; zero behavioral checks. v38 worker passed 90/90, exited naturally and had zero delayed faults, but its controller expected 85 names because the fixture itself emits five compile checks. These are controller/setup failures, not product RED. Both approvals consumed.
- v39 PR4-BASE-10 visible Event Detail: **90/90 exact PASS**, including five fixture-owned compile checks, worker exit 0, controller `StopReason=null`, natural Excel exit 2.52 seconds after stdout quiet, settings restored, 539 pins preserved, no intervention or process termination, delayed audit zero faults. Reviewed default, multiline-bottom, multiline-right, maximized and restored captures show the selected detail, readable long/multiline text, reachable last line and horizontal `END` marker, and reachable Close. Evidence: `reports/runtime/viewer-detail-v39-pr4-base-10-20261112` and its `ReportRoot`.
- Read-only next-gate map: latest historical Viewer published-read report has 80 unique checks (2026-10-21): five fixture compile checks, one recording-probe setup check, 36 PublishedRead and 38 Recording checks. This appears to be a combined recording/published-read selector; do not assume `-CheckViewerPublishedRead` alone emits 80. Groups has 82 (2026-09-14), refresh 16 (2026-09-14). These are older-candidate results, not current-package qualification. The current worker supports a compiled published-read gate, but its exact selector and emitted identity set must be established before pinning.
- Plan 022 / Controls 1.745 are pushed on docs PR5. These standalone gates do not establish complete Event Viewer, 4be-A, human Receiving B0 or Release 1 acceptance.

## Do Not Repeat

Do not rerun v32-v39, count v37/v38 as product RED or qualified gates, change a pinned run after approval, or merge the older `soetrain` docs working branch. Avoid translating UTC cutoffs without checking Pacific time.

## Assumptions to Re-verify

Recheck PR heads, clean worker, Excel closure, package hashes, desktop/input, and whether the candidate has changed before the next packaged gate.

## Open questions and blockers

Current-package Viewer published-read/groups/refresh, guide transfer, other Admin/Operations, retained formatting, broad Inventory, human Receiving B0, live role, full chain and visible operator acceptance. Execute slowdown and the Microsoft-reported Excel exit defect remain separate.

## Immediate next action

Determine the exact current-worker selector for the combined recording/published-read gate, pin its emitted identity set, and prepare one packaged gate on the unchanged five XLAMs; request bounded approval only after checks-only passes.

## Critical references

Architecture v4.11 D18 and Event Detail scrolling decision; Plan 022 PR4-BASE-10 checkpoint; Controls 1.745; code PR14; docs PR5; `tests/tooling/Test-Slice4beConfigCommands.ps1`, `Slice4beViewerEventDetail.ps1`, `Slice4beViewerPublishedRead.ps1`; v36/v39 evidence folders above.
