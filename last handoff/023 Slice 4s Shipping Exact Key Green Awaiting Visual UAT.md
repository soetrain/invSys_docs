# Goal and release outcome

Complete the visible Release 1 Receiving → Production → Boxing → Shipping
workflow under specification v4.11. Plan 022 Slice 4s advances D14 by making
Shipping Add reserve the selected immutable `System_Key`, while Box Designer
shows only active exact component entities.

# Current verified state

- Code: `main` at pushed commit `c4ff28c` (`Fix Shipping exact-key reservation`).
- Docs: `main` at `0e8be1e` before this handoff commit; this session updates the
  current plan, controls inventory, handoff, and pointer.
- Active slice: Plan 022 Slice 4s, automated GREEN; visible `WHT7025AE` retest
  remains. Last verified 2026-08-20.
- Rebuilt/deployed `invSys.Core.xlam` and `invSys.Operations.xlam` plus manifest.
- Excel is fully closed. Last verified 2026-08-20.
- Preserved unstaged user changes: code `AGENTS.md`; the NAS/server note at the
  top of `invSys-Design-v4.11.md`.

# Decisions and constraints

- Shipping Add must never require managed `ROW`; its current-schema local
  reserve apply uses the selected string `System_Key`.
- Box Designer omits nonpositive balances and duplicate projections of the same
  key. Distinct positive entities remain distinct even when display fields match.
- Required persistence remains in place. Shipping/Boxing actions use the shared
  quiet boundary, which now also hides/restores Excel's status bar.
- Native Saving-window suppression is not considered visibly accepted until the
  user repeats the actions in the exact NAS session.

# Evidence and traceability

- Symptom: Shipping Add reported missing `TOTAL INV/SHIPMENTS/ROW` → root cause:
  `ApplyShipmentDeltasLocal` lacked a no-`ROW` exact-key branch → D14 → protected
  by `TestShippingAdd_CurrentSchemaReservesBySystemKey` through the same public
  `ShipmentsFormCommitLine` action.
- Focused Slice 4s RED: 1/4 then 4/6; GREEN: 6/6.
- Adjacent form-action range: 3/3; packaged XLAM: 74/74; live role workflows:
  46/46; ordered Release 1 full chain: 30/30.
- Shipping/Boxing: 11/11; final controls: 12/12; workflow readiness: 18/18.
- Static maintenance: 19/19; 150 components, 4,702 procedures, 965 candidates,
  8 literal `Application.Run` targets, 47 unresolved expressions, 184
  duplicate-body groups.
- Generic visible package inspection passed 32/34; its two failures are stale
  expectations for retired `AggregateBoxBOM_Log` and `AggregatePackages_Log`
  worksheets, not active form/action failures.

# Do Not Repeat

- Do not satisfy Shipping by restoring or relabeling numeric `ROW` identity.
- Do not collapse distinct positive entities into one Box Designer row unless a
  future allocation design preserves every exact key.
- Do not remove required workbook saves solely to suppress visible notifications.
- Do not revive retired Shipping support worksheets to satisfy the generic
  visible inspector.

# Assumptions to Re-verify

- The exact `WHT7025AE` component picker omits all zero-balance and same-key
  duplicate rows after restart.
- Hiding the Excel status bar suppresses the reported Saving notifications on
  this Office/NAS environment.
- Shipping Add visibly reserves the selected box and refreshes the list without
  a `ROW` error.

# Open questions and blockers

- Visible UAT is still required for Box Designer component rows, Shipping Add,
  and Saving notifications during Confirm Writes/Dispositions, Save Box,
  Make/Unmake, Add, and Shipments Sent.

# Immediate next action

Open Excel normally, select/sign into `WHT7025AE`, then retest Box Designer
Component inventory and one exact-key Shipping Add before continuing Shipments
Sent.

# Critical references

- `src/Shipping/Forms/frmShipmentsTally.frm:CommitCurrentLine`
- `src/Shipping/Modules/modTS_Shipments.bas:ApplyShipmentDeltasLocal`
- `src/Shipping/Modules/modTS_Shipments.bas:ApplyShipmentDeltasBySystemKey`
- `src/Shipping/Modules/modTS_Shipments.bas:BuildShippingInventoryPickerItems`
- `src/Shipping/Modules/modTS_Shipments.bas:BuildCanonicalRuntimeInventoryPickerItems`
- `src/Shipping/Modules/modBoxingService.bas:SaveBoxDesign`
- `src/Shipping/Modules/modBoxingService.bas:PostBoxMakerAction`
- `src/Core/Modules/modUiQuiet.bas`
- `tests/tooling/Test-Plan022Slice4sShippingExactKey.ps1`
- `tests/integration/plan022_slice4s_shipping_exact_key_results.md`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
