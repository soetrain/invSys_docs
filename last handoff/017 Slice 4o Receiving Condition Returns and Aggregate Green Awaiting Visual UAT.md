# Goal and release outcome

Complete Release 1 operator workflows against the dedicated NAS warehouse; Slice 4o advances D14 Receiving identity by establishing condition at receipt, supporting inbound returns, and rebuilding the complete aggregate projection.

# Current verified state

- Code: `main` at `e549328` (`Add receiving condition returns and aggregate repair`), pushed 2026-08-19.
- Docs: `main` at `9e89372` (`Document receiving condition and returns contract`), pushed 2026-08-19.
- Active slice: Plan 022 Slice 4o, automated GREEN; visible `WHT7025AE` UAT remains.
- Rebuilt and committed `invSys.Core.xlam`, `invSys.Operations.xlam`, and `invSys.Admin.xlam` plus the deployment manifest.
- Excel is fully closed. Last verified 2026-08-19.
- Preserved unstaged user changes: code `AGENTS.md`; the NAS/server note at the top of `invSys-Design-v4.11.md`.

# Decisions and constraints

- Receiving, not Inventory Viewer, establishes physical `Condition`; R1 choices are `GOOD`, `BAD`, `DAMAGED`, `EXPIRED`, and `REJECTED`.
- `System_Key` is durable entity identity. `Lot` is independent provenance. Operational hold/quarantine state remains conceptually separate from physical condition.
- Mixed-condition receipt lines create separately keyed inventory entities.
- Receiving **Returns** is for inbound returned goods and uses the authorized `RECEIVE` boundary; outbound vendor returns are outside this slice.
- Demo Inventory has no Cancel button; closing the form without an action is silent.

# Evidence and traceability

- Focused source RED 0/5 -> GREEN 5/5: `tests/integration/plan022_slice4o_receiving_results.md`.
- Same-form public-action RED 0/2 -> GREEN 4/4: tests 288-291 in `TestReceivingStabilization.bas`.
- Packaged Slice 4o actions 5/5, including a BAD-condition inbound return and silent Demo Inventory close.
- Original Aggregate Received symptom -> MSForms ListBox was populated with 11 columns although its limit is 10 -> complete 10-column projection and rebuild from every tally row -> focused refresh/action and packaged tests.
- Affected bootstrap/migration workbook-surface range 28/28; packaged XLAM 74/74; packaged RibbonX 142/142.
- Ordered deployed Release 1 chain 30/30, including restart and reconciliation.
- Static baseline deterministic 19/19; 150 components, 4,680 procedures, 961 scanner candidates, 963 reviewed candidates, 24 oversized-module ratchets. Slice 12 metric ratchets 11/11.

# Do Not Repeat

- Do not restore the 11-column MSForms aggregate/history list; backing tables may retain more evidence columns, but visible list projections must stay at ten or fewer.
- Do not make Inventory Viewer an editor for condition.
- Do not treat condition, lot, location, quantity, or SKU as durable identity.
- The first full-chain rerun passed 28 business checks but its runtime extractor left a hidden Excel process; clearing that exact orphan and rerunning produced the clean 30/30 gate.

# Assumptions to Re-verify

- Visible resizing and header alignment of the expanded Receiving lists have automated geometry coverage but still need operator inspection on the target display.
- The built and committed package set must be the package set Excel loads during NAS UAT.

# Open questions and blockers

- Visible UAT has not yet confirmed mixed GOOD/BAD lines, every aggregate row/total, and one inbound return against `WHT7025AE`.
- Whether R1 needs an operator-facing `InventoryState` hold/quarantine workflow remains unresolved; it must remain separate from `Condition` if added.

# Immediate next action

Open the rebuilt package set and run one same-reference mixed-condition receipt plus one inbound return in `WHT7025AE`, then compare Received Tally, Aggregate Received, receipt history, and Inventory Viewer after Confirm Writes/refresh.

# Critical references

- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4o
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, section 7
- `src/Receiving/Forms/frmReceiving.frm`
- `src/Receiving/Modules/modTS_Received.bas`
- `src/Receiving/Modules/modReceivingPostingService.bas`
- `src/Core/Modules/modRoleWorkbookSurfaces.bas`
- `tests/tooling/Test-Plan022Slice4oReceiving.ps1`
- `tools/validate_plan022_slice4o_packaged.ps1`
- `tests/integration/slice14_results.md`
