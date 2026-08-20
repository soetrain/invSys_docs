# Goal and release outcome

Complete Release 1 Receiving -> Production -> Boxing -> Shipping acceptance
under specification v4.11. Plan 022 Slice 4t advances D13/D14 by making the
public Shipments Sent action refresh and render the canonical post-send balance.

# Current verified state

- Code: `main` at pushed commit `9c5d7e2` (`Refresh Shipping inventory after send`).
- Docs: `main` at pushed commit `29cb517` before this handoff/pointer commit.
- Active slice: Plan 022 Slice 4t, automated GREEN; exact `WHT7025AE` visual
  retest remains. Last verified 2026-08-20.
- Rebuilt/deployed `invSys.Core.xlam` and `invSys.Operations.xlam` plus manifest.
- Excel is fully closed. Last verified 2026-08-20.
- Preserved unstaged user changes: code `AGENTS.md`; the NAS/server note at the
  top of `invSys-Design-v4.11.md`.

# Decisions and constraints

- `modProcessor.RunBatch` owns canonical apply and snapshot generation. Shipping
  must not publish that snapshot a second time after the batch.
- A successful Shipments Sent action refreshes the captured operator workbook,
  reloads shippables, then derives NAS Inv, Projected Inv, and Locked.
- The processor's three durability saves remain required. Removing the redundant
  Shipping publication should reduce notices/time, but visible suppression is
  not accepted until the operator recounts native Saving notices.
- `System_Key` remains the exact shipment/package identity; no `ROW` fallback is
  permitted on current schemas.

# Evidence and traceability

- Symptom: canonical Viewer changed 100 -> 95 after five packages, while the
  Shipping list remained 100/100/0 -> root cause: no post-batch operator
  read-model refresh plus cached `mShippables` render -> Slice 4t contract ->
  protected by the public form-action live assertion and focused ordering test.
- Timing evidence: reported action about 84 seconds; batch about 58 seconds;
  twenty Saving notices. The removed duplicate publication occupied the
  post-batch boundary; exact new timing remains a visual measurement.
- Focused RED: 0/4; focused GREEN: 4/4.
- Packaged XLAM: 74/74; live role workflow: 47/47; ordered full chain: 30/30.
- Shipping/Boxing: 11/11; final controls: 12/12; workflow readiness: 18/18.
- Static maintenance: 19/19; 150 components, 4,702 procedures, 964 candidates,
  8 literal `Application.Run` targets, 47 unresolved expressions, and 184
  duplicate-body groups.
- One initial ordered-chain run completed all business assertions but its final
  read-only extractor attached to an orphaned validation Excel process. After
  removing that test-owned process, the clean rerun passed 30/30.

# Do Not Repeat

- Do not add `PublishInventorySnapshotBridge` back after `RunBatch`; it duplicates
  the processor's snapshot boundary.
- Do not remove the processor's three saves merely to hide native notices.
- Do not accept Inventory Viewer alone as proof; the still-open Shipping form is
  the contract under test.
- Do not hard-code a test balance of 15. The live assertion uses starting
  quantity minus shipped quantity so both normal and ordered fixtures pass.

# Assumptions to Re-verify

- The reported existing box entity remains at canonical quantity 95 when Excel
  next opens.
- The rebuilt package is the package loaded by the user's normal Excel startup.
- Removal of the duplicate publication measurably reduces elapsed time and/or
  native Saving notices on the NAS connection.

# Open questions and blockers

- Visible `WHT7025AE` proof remains for NAS Inv 95, Projected Inv 95, Locked 0
  immediately on Shipping load and after the next Shipments Sent action.
- Exact Saving-notice counts and action timing remain unaccepted.
- Box Designer's filtered component list still needs a visible confirmation that
  zero balances and same-key duplicate projections are absent.

# Immediate next action

Open Excel normally, select/sign into `WHT7025AE`, open Shipping, and verify the
completed box row is 95/95/0 before timing one small Shipments Sent action and
counting its native Saving notices.

# Critical references

- `src/Shipping/Forms/frmShipmentsTally.frm:mBtnSend_Click`
- `src/Shipping/Forms/frmShipmentsTally.frm:TestRunShipmentsSentActionForWorkbook`
- `src/Shipping/Modules/modTS_Shipments.bas:RunShippingRuntimeQueueRefresh`
- `src/Shipping/Modules/modTS_Shipments.bas:ShipmentsFormRefreshReadModelForWorkbook`
- `src/Core/Modules/modProcessor.bas:RunBatch`
- `tools/validate_phase6_live_role_workflows.ps1`
- `tests/tooling/Test-Plan022Slice4tShippingPostSend.ps1`
- `tests/integration/plan022_slice4t_shipping_post_send_results.md`
- `tests/integration/slice14_results.md`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
