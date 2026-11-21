# Handoff 223: PR4-BASE-02 Inventory gate evidence

Date: 2026-11-08 Pacific. Predecessor: handoff 222.

## Goal and release outcome

Complete invSys Release 1 Slice 4be-A under Architecture v4.11 and Plan 022. The v16 packaged edit-creation/Config gate advances D13 evidence but does not establish broad Inventory GREEN or release acceptance.

## Current verified state

Last verified: 2026-11-08 Pacific. Code branch work/d8-inv-edit-receipts is at ebb3deb43ead26e9d865742e68e8641ee3a1ee41 (draft PR4). Docs branch work/perplexity-handoff218 was at 9123ee5252bd8d2e1b1d410a142c5c744730a309 before this handoff (draft PR5). Plan 022 and Controls 1.733 hold the refreshed checkpoint. No VBA, XLAM, architecture or product contract changed. The main code checkout has Justin's unrelated modified AGENTS.md. The docs checkout has pre-existing unrelated changes: deletion of expert guidance docs/024 Slice 4be Delivery Overrun and Completion Review.md; edit to last handoff/067 Partial Goal Action Path NAS Contract.md; untracked WSAIFabricSvc with a reboot for one AB run.txt, expert guidance docs/023 Slice 4be Critique.md, 024 Slice 4be REPORT Delivery Overrun and Completion Review.md, and 025 GUIDANCE Action Path contract vs expectation conflict.md. Preserve them. At the v16 delayed audit Excel was closed, settings restored and 538 frozen inputs preserved; recheck before another test or build.

## Decisions and constraints

Justin approved separate test-only PR4-BASE-02 manifest SHA256 F2DE6B82..., re-pinning only tests/tooling/Slice4beInventoryEditCanonicalSave.ps1 to SHA256 E0428CB0.... PR4-BASE-01 SHA256 6FDE11C2... and all five XLAM pins remain unchanged. The probe selects the loaded Domain package and stops clearly if it cannot be inspected; no fallback to the old Admin path. v16 stands alone and is not added to v13.1's failed run. ControlIds and wording did not change.

## Evidence and traceability

- v14 ran BeforeCatalogSaveThrows and AfterCatalogSaveThrows separately and repeated v13.1's same two failed checks: 17 PASS, 2 FAIL, with 9 and 14 missing respectively. The combined session was not the cause.
- The test's source-tree check chose the old Admin save path because the main checkout lacks the Domain module. v15 used the loaded Domain package and passed 28/28 and 33/33. PR4 commit e2c8a06f changes only this test-tooling file.
- v16 on PR4-BASE-02 passed EditCreation.All 545/545 (including 140 QuantityPhase), ConfigPreservation.All 50/50, MissingDefault 54/54 and RetainedShared 123/123. All 772 identities matched exactly, four workers exited 0, Excel exited naturally, settings/inputs were preserved and the delayed audit found zero Excel faults. Run 01 AfterQuit-to-exit was 14.95 s; Config selectors have only informational stdout-quiet-to-exit times. The code Inventory audit holds details; local runtime reports remain ignored.

## Do Not Repeat

Do not infer a product save defect from v13.1/v14, rerun those diagnostics to explain the known probe-owner fault, modify frozen manifests, or combine v16 with v13.1. Do not claim the Excel native-exit defect is fixed. Justin submitted a Microsoft report; no follow-up dialogue is open.

## Assumptions to Re-verify

Before runtime work, verify Excel closure, desktop input, package/manifest pins, current branch and PR heads, settings restoration and no unfinished worker. The slowdown across Execute calls remains unexplained.

## Open questions and blockers

Complete retained-formatting caller coverage remains open despite RetainedShared 123/123. Other Inventory selectors, broad Inventory GREEN, recording and guides, early Receiving B0, live-role, restart, full Release 1 chain, final compile/layout/static and visible-operator evidence, and 4be-A acceptance remain open. The Excel native-exit defect remains unresolved. Separate 4be-B replay acceptance is outside this gate.

## Immediate next action

Map unqualified Inventory selectors and pin exact expected identities for the first focused packaged D13 regression before another Excel dispatch.

## Critical references

Architecture v4.11 D13, D8-INV-EDIT-01/02 and D8-INV-CONFIG-01; Plan 022; Controls 1.733; tests/tooling/Slice4beInventoryEditCanonicalSave.ps1; PR4-BASE-01/02 manifests; reports/runtime/inventory-creation-save-v14-20261108/, inventory-creation-save-v15-20261108/ and inventory-gate-v16-pr4-base-02-20261108/; draft code PR4 and docs PR5.
