# Handoff 234: Viewer published-read break and visibility comparator

Date: 2026-11-13 Pacific. Predecessor: handoff 233.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. The current five-XLAM Viewer published-read/recording gate stopped before behavioral checks; 4be-A acceptance remains open.

## Current verified state

Last verified 2026-11-13: code `main` at `6d247543` has Justin's unrelated `AGENTS.md` edit. Draft code PR14 head is `b8f94241`. The execution worker was restored clean to `work/d8-inv-edit-receipts` at `550d9ff0` after v40; detach at `b8f94241` again only for an approved v41 run. Docs branch is `work/perplexity-handoff218` (pre-handoff head `9fd9196`); unrelated docs changes are untouched. Excel and the v40 controller are closed. v36 Settings 202/202 and v39 Event Detail 90/90 remain separate qualified gates.

## Decisions and constraints

- Product VBA, Architecture v4.11, five XLAMs, ControlIds and displayed wording are unchanged. v40/v41 are ignored runtime controllers only.
- The owner's active shift ends by 2026-11-13 **21:45 UTC / 1:45 PM Pacific**. v41's prepared cutoff is 20:45 UTC; approval is pending and elapsed time is not authorization.
- v40's one-run authorization is consumed. No v40 retry or assisted Excel shutdown is authorized. Any v41 Excel run needs its own owner authorization naming the manifest, expected-check and launcher hashes. Operator interaction with v40's VBA window is being confirmed; treat attendance as unknown until answered.

## Evidence and traceability

- v40 checks-only: 539 PR4-BASE-10 pins and 80 exact expected identities; no Excel. Approved visible run: five compiles, recording probe and two publication checks PASS, then VBA break at `modInventoryViewer.OpenInventoryViewer` -> `mInventoryViewer.SetWarehouse`. No modal; captured GDI/USER near 2,100. Worker ultimately reported **8 PASS / 1 fixture exception / 72 missing checks** and exit 1. Controller exact-identity gate failed; both native exits verified, invSys test settings restored, 539 pins preserved, no process termination, delayed audit zero Excel faults. This is not product RED or qualified GREEN. Evidence: `reports/runtime/viewer-published-recording-v40-pr4-base-10-20261113/` and its run-S1 capture/image.
- Historical 2026-10-20 80/80 recording gate started Excel hidden; its fixture revealed Excel immediately before Viewer open. v40 started visible. The separate v41 comparator changes only that initial-host visibility timing. v41 parser and checks-only pass (539 pins, 80 identities); no v41 Excel dispatch yet. Evidence/instructions: `reports/runtime/viewer-published-recording-hidden-v41-pr4-base-10-20261113/`.
- Read-only next-gate map: groups/publication has 82 historical checks; refresh has 16. Current fixture guards reject a command-line compile switch for those selectors. See v40 folder `next-viewer-gate-map.md`. These are older-candidate identities, not current-package qualification.

## Do Not Repeat

Do not rerun v40, count its exception as behavioral RED, call it unattended without the owner's answer, or treat v41 checks-only as runtime GREEN. Do not add the compile switch to groups/refresh controllers. Preserve the owner's unrelated edits.

## Assumptions to Re-verify

Recheck v41 approval, start cutoff, worker cleanliness, Excel closure, package hashes, desktop input and the v40 interaction answer before dispatch. The exact VBA runtime error and cause of break mode remain unknown.

## Open questions and blockers

Current-package Viewer published-read/recording, groups/refresh, guide transfer, other Admin/Operations coverage, retained formatting, broad Inventory, human Receiving B0, live-role, full-chain and visible operator acceptance remain open. Execute slowdown and the Microsoft-reported Excel exit defect remain separate.

## Immediate next action

If the owner approves v41 within its cutoff, confirm a clean worker, detach at `b8f94241`, create its single-use authorization, repeat checks-only and execute the hidden-start comparator once; then restore the worker branch and classify the result before any product change.

## Critical references

Architecture v4.11 D18; Plan 022 PR4-BASE-10 checkpoint; Controls 1.746; code PR14; docs PR5; `tests/tooling/Test-Slice4beConfigCommands.ps1`, `Slice4beViewerPublishedRead.ps1`, `Slice4beActionRecording.ps1`; v40/v41 runtime folders above.
