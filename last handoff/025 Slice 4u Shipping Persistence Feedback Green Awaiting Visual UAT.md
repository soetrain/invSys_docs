# Goal and release outcome

Complete Release 1 Receiving -> Production -> Boxing -> Shipping acceptance
under specification v4.11. Plan 022 Slice 4u advances D13/D14 by consolidating
Shipping persistence feedback without weakening durable NAS writes.

# Current verified state

- Code: `main` at pushed-candidate commit `ac726bd` (`Consolidate Shipping
  persistence feedback`).
- Docs: `main` at commit `16c90e9` before this handoff/pointer commit.
- Active slice: Plan 022 Slice 4u, automated GREEN; the new persistence-summary
  line needs one visible `WHT7025AE` retest. Last verified 2026-08-20.
- Rebuilt `invSys.Core.xlam` and `invSys.Operations.xlam` plus the manifest in
  `deploy/current`.
- Excel is fully closed. Last verified 2026-08-20.
- Preserved unstaged user changes: code `AGENTS.md`; the NAS/server note at the
  top of `invSys-Design-v4.11.md`.

# Decisions and constraints

- A one-row Shipping Add retains two logical server saves: warehouse inbox and
  Shipping reservation ledger. Shipments Sent also retains the processor's
  three durability saves.
- Excel-native Saving progress windows cannot be moved into the VBA form.
  invSys now reports one authoritative `Persistence summary:` line in the
  existing Shipping status/message output.
- Multi-row reservation upserts open and save the reservation ledger once per
  action, not once per selected row.
- `System_Key` remains the exact shipment/package identity; no current-schema
  `ROW` fallback is permitted.

# Evidence and traceability

- User-visible Slice 4t acceptance: the still-open Shipping list matched Viewer
  after six packages shipped. Total was about 65 seconds, processor batch about
  49 seconds, and read-model refresh 836 milliseconds.
- Remaining symptom: one-row Add displayed four Excel-native Saving notices ->
  required inbox and reservation durability plus Excel-owned progress UI ->
  Slice 4u consolidated-feedback contract -> protected by the real Add and
  Shipments Sent callback test.
- Focused RED: 1/4; focused GREEN: 4/4.
- Packaged XLAM: 74/74; live role workflow: 47/47; ordered full chain: 30/30.
- Shipping/Boxing: 11/11; final controls: 12/12; workflow readiness: 18/18.
- Reviewed cleanup: 11/11; 150 components, 4,702 procedures, 964 candidates,
  8 literal `Application.Run` targets, 47 unresolved expressions, and 184
  duplicate-body groups.
- Two ordered-chain attempts passed every business assertion but ended on the
  known test-owned orphaned Excel automation process; after removing that
  process, the clean final rerun passed 30/30.

# Do Not Repeat

- Do not remove inbox, reservation-ledger, inventory, outbox, or inbox-processor
  saves merely to suppress Excel-native progress UI.
- Do not claim `Application.DisplayAlerts` or `DisplayStatusBar` can suppress or
  reparent every native Excel save-progress window; the quiet boundary already
  uses both.
- Do not restore per-row reservation workbook open/save cycles.
- Do not add a second snapshot publication after `modProcessor.RunBatch`.

# Assumptions to Re-verify

- The account-scoped Excel add-in registration still loads packages from
  `deploy/current`.
- The new `Persistence summary:` line is readable in the visible Shipping
  status area and Shipments Sent dialog at the user's display scale.
- Excel may still show native Saving notices for the required two Add writes and
  the required Shipments Sent writes.

# Open questions and blockers

- Visible confirmation remains for the new persistence summary after Add and
  Shipments Sent, with native Saving notices counted separately.
- The roughly 49-second processor batch remains the principal Shipping latency;
  Slice 4u does not change that durability path.
- Box Designer's filtered component list still needs visible confirmation that
  zero balances and same-key duplicate projections are absent.

# Immediate next action

Open Excel normally, select/sign into `WHT7025AE`, run one Shipping Add and one
Shipments Sent action, and confirm each ends with one `Persistence summary:`
line while separately recording any Excel-native Saving notices.

# Critical references

- `src/Shipping/Forms/frmShipmentsTally.frm:CommitCurrentLine`
- `src/Shipping/Forms/frmShipmentsTally.frm:mBtnSend_Click`
- `src/Shipping/Modules/modTS_Shipments.bas:ShipmentsFormCommitLine`
- `src/Shipping/Modules/modTS_Shipments.bas:ShipmentsFormRunToShipmentsRows`
- `src/Shipping/Modules/modTS_Shipments.bas:ShipmentsFormRunShipmentsSentRows`
- `src/Shipping/Modules/modTS_Shipments.bas:UpsertShippingReservationRows`
- `src/Shipping/Modules/modTS_Shipments.bas:AppendShippingPersistenceSummary`
- `tests/tooling/Test-Plan022Slice4uShippingPersistenceFeedback.ps1`
- `tests/integration/plan022_slice4u_shipping_persistence_feedback_results.md`
- `tests/integration/slice14_results.md`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
