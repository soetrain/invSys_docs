# Goal and release outcome

Complete Release 1 operator workflows against the dedicated NAS warehouse; Slice 4q corrects Returns so `RETURN` and `DUMP` deplete exact existing inventory entities instead of creating inbound inventory.

# Current verified state

- Code: `main` at `9da376f` (`Make returns deplete exact inventory`), pushed 2026-08-19.
- Docs: `main` at `3b1c945` (`Define outbound return and dump controls`), pushed 2026-08-19 before this handoff commit.
- Active slice: Plan 022 Slice 4q, automated GREEN; visible `WHT7025AE` UAT remains.
- Rebuilt and committed `invSys.Core.xlam`, `invSys.Inventory.Domain.xlam`, and `invSys.Operations.xlam` plus the deployment manifest.
- Excel is fully closed. Last verified 2026-08-19.
- Preserved unstaged user changes: code `AGENTS.md`; the NAS/server note at the top of `invSys-Design-v4.11.md`.

# Decisions and constraints

- The Returns page is outbound disposition. `RETURN` means goods leave for an external party; `DUMP` means goods are discarded. Both reduce on-hand inventory.
- The operator enters a positive quantity. Domain audit rows apply a negative delta under distinct `RETURN`/`DUMP` event types and the existing `RECEIVE_POST` capability.
- Disposition preserves the exact existing `System_Key`, SKU, location, lot, and Condition. It creates no new inventory key and rejects exact-entity overdraw even when another entity has the same SKU.
- A visible aggregate choice is deterministically expanded into separately keyed allocations without crossing item/location/lot/Condition boundaries.
- On Returns, location, lot, and Condition are selection-backed and locked. The required Disposition selector defaults to `RETURN`; the actions read **Add Disposition** and **Confirm Dispositions**.
- Historical wrongly applied inbound-return rows are audit evidence and were not silently rewritten. In the reported packing-tape state, applying a new 50-unit RETURN to the displayed 150 should produce 100; a clean 100-unit starting entity minus 50 produces 50.

# Evidence and traceability

- D13 RED: focused static contract 0/6; same-handler form action returned 0; Domain disposition test returned 0.
- Focused GREEN: aligned VBA range 290-296 is 7/7; static Slice 4q contract 6/6; packaged Returns action 5/5.
- Inventory Viewer action loaded three snapshot levels, reused one form generation, filtered locally, and left its snapshot byte-for-byte unchanged.
- Packaged XLAM 74/74; RibbonX 142/142; live role workflow 46/46; clean ordered Release 1 chain 30/30.
- Static maintenance: Slice 3 19/19, 150 components, 4,698 procedures, 965 scanner candidates; Slice 12 11/11 with a protected Slice 4q procedure ceiling and no component/candidate/duplicate-body/late-binding regression.

# Do Not Repeat

- Do not restore inbound `RECEIVE` semantics for the Returns page or generate a replacement `System_Key`.
- Do not validate a disposition against SKU-level total inventory; validate each allocated exact entity.
- Do not let Return/Dump edit source location, lot, or Condition.
- Do not silently delete or offset the historical incorrect inbound event in `WHT7025AE`; use clean reseed/UAT data or an explicit audited correction decision.

# Assumptions to Re-verify

- Excel will load the rebuilt add-ins from `deploy/current` on the next start.
- The current `WHT7025AE` packing-tape projection still includes the historical incorrect 50-unit inbound return and therefore starts at 150 DAMAGED units.

# Open questions and blockers

- Visible UAT has not yet confirmed the selector, RETURN/DUMP depletion, Viewer refresh, exact Condition separation, and overdraw message against `WHT7025AE`.
- The complete Release 1 user-side workflow checkpoint remains open after this focused disposition retest.

# Immediate next action

Start Excel, open `WHT7025AE` Returns, RETURN 50 from the displayed 150-unit DAMAGED packing-tape row, confirm Refresh and Inventory Viewer both show 100, then DUMP a smaller quantity and verify a second exact decrease.

# Critical references

- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`, Receiving condition and inventory-disposition rule
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4q and UAT step 14
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, section 7
- `src/Receiving/Forms/frmReceiving.frm`
- `src/Receiving/Modules/modTS_Received.bas`
- `src/Receiving/Modules/modReceivingPostingService.bas`
- `src/Core/Modules/modRoleEventWriter.bas`
- `src/Core/Modules/modProcessor.bas`
- `src/InventoryDomain/Modules/modInventoryApply.bas`
- `tests/tooling/Test-Plan022Slice4qDisposition.ps1`
- `tests/unit/TestReceivingStabilization.bas:TestReceivingReturns_StagesExistingDispositionIdentityThroughFormAction`
- `tests/unit/TestPhase6CoreSurfaces.bas:TestInventoryDisposition_ReturnAndDumpDepleteExactSystemKey`
- `tools/validate_plan022_slice4o_packaged.ps1`
