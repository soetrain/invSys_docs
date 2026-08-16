# Slice 4j Automated Green Awaiting Dedicated NAS UAT

## 1. Goal and release outcome

Complete Plan 022 Slice 4j so the five-package Release 1 build supports a testable Receiving -> Production Run - List -> Box Maker -> Shipping workflow under v4.11 D12-D14. Automated packaged acceptance is complete; the remaining release evidence is the visible dedicated-NAS operator checkpoint.

## 2. Current verified state

- Last verified: 2026-08-16.
- Code: `main`, commit `2af50c4` (`Complete R1 workflow control readiness`).
- Docs: `main`, commit `19911c5` (`Document Slice 4j workflow readiness`).
- Active slice: Plan 022 Slice 4j, automated gates complete.
- Excel was closed after all builds and validators.
- User-owned unstaged changes remain in code `AGENTS.md` and docs `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`; do not overwrite or stage them without review.

## 3. Decisions and constraints

- Ordinary Station is the Windows computer name; Test Environment Setup remains a separate disposable-fixture utility.
- Exact saved role names are `<WarehouseId>.Receiving.Operator.xlsm`, `<WarehouseId>.Production.Operator.xlsm`, and `<WarehouseId>.Shipping.Operator.xlsm`.
- Box Designer/Box Maker alternatives retain compatible `v1`-style storage labels, but visible wording is Alternative.
- Physical inventory, Shipping BOM packages/components, and boxing events preserve immutable string `System_Key`; Production Run - Tree remains experimental.
- The complete demo kit contains 24 new `Condition=GOOD` entities, including carton, divider, label, tape, and void fill.

## 4. Evidence and traceability

- Focused Slice 4j: 18/18 GREEN after an initial 1/13 RED.
- Five-package build succeeded; packaged XLAM 74/74; packaged Ribbon 142/142.
- Production scaling, Receiving durability/search/header alignment, and Shipping resize/identity packaged launcher probes each pass 1/1.
- Public Admin seed callback exposes 24/24 unique entities through canonical inventory, snapshot, and saved Receiving refresh.
- Ordered isolated full chain is 30/30 through fresh warehouse, seed, Receiving, two Production batches, Box Maker, Shipments Sent, restart, replay, and reconciliation.
- Slice 5 behavior locks 13/13; Slice 11 Shipping/Boxing 11/11; Slice 13 Operations cutover 14/14; Ribbon generation 46/46; deterministic static baseline 19/19.
- Static baseline: 150 components, 4,614 procedures, literal `Application.Run` 9, unresolved expressions 48, duplicate bodies 184.

## 5. Do Not Repeat

- The apparent Box Maker processor hang was a hidden VBA compile dialog caused by the nonexistent `SyncLocalStagedInboxSystemKeys` call. The supported Core API is `SyncLocalStagedInboxRows`; this is now guarded by the focused test.
- Do not treat Receiving's former duplicate inventory list as an inventory viewer. Its top list is Receiving Entries History; Inventory Viewer is the Operations overview.
- Do not add compatibility mappings back to `ROW` in the reachable boxing path.

## 6. Assumptions to Re-verify

- Verify the dedicated warehouse accepts the current computer-name Station and that the signed-in user's capability rows use the same station scope.
- Recheck visible maximize/restore/grow/shrink behavior at the operator's display scaling even though packaged geometry is GREEN.
- Confirm the 24-row seed is visible in Inventory Viewer on the dedicated warehouse.

## 7. Open questions and blockers

- No automated blocker remains.
- Visible dedicated-NAS acceptance is outstanding for Receiving Location/Lot, Production scales `0.001%`/`100%`/`1000%`, Box Designer/Box Maker resizing, and the complete seed-to-ship workflow.
- Static maintenance still reports historical `ROW` references outside the reviewed Release 1 packaged authority path; remove only through a separately scoped D13 reachability slice.

## 8. Immediate next action

Run Plan 022 section 6's single dedicated-NAS operator checkpoint beginning with the 24-row Seed Demo Inventory action and Inventory Viewer refresh.

## 9. Critical references

- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4j and section 6
- `tests/tooling/Test-Plan022Slice4jWorkflowReadiness.ps1`
- `tests/integration/admin_seed_callback_green_results.md`
- `tests/integration/slice14_results.md`
- `reports/static-baseline/`
- `src/Core/Modules/modStationIdentity.bas`
- `src/Core/Modules/modRoleWorkbookSurfaces.bas`
- `src/Receiving/Forms/frmReceiving.frm`
- `src/Production/Forms/frmProduction.frm`
- `src/Shipping/Forms/frmShipmentsTally.frm`
- `src/Shipping/Modules/modTS_Shipments.bas`
