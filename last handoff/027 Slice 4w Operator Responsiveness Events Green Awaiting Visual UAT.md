# Goal and release outcome

Complete the Release 1 operator workflow by making remote connection/persistence waits understandable, making concatenated Receiving references readable, and adding read-only cross-role event visibility to Inventory Viewer.

# Current verified state

- Code branch `main`, implementation commit `ab53da0` (`Add operator event visibility and progress feedback`).
- Documentation branch `main`, contract commit `822abe6` (`Document operator responsiveness and event visibility`).
- Active slice: Plan 022 Slice 4w — operator responsiveness and read-only event visibility.
- Focused RED/GREEN, package build, packaged workflows, Viewer form action, full chain, and maintenance gates are complete.
- Excel is fully closed. Last verified: 2026-08-20.
- The user's pre-existing code `AGENTS.md` and documentation `invSys-Design-v4.11.md` edits remain uncommitted and were not modified by this slice.

# Decisions and constraints

- MSForms ListBox rows cannot have variable height or wrapped cells. `frmReceiving` uses a locked multiline **Selected references** box for the complete selected aggregate reference.
- Windows SMB authentication remains a synchronous `WNetAddConnection2` call. invSys paints progress and yields before it; Windows may still mark Excel busy until the network call returns.
- Excel-native Saving windows cannot be reparented into a VBA UserForm. Receiving/Returns, Production, and Shipping paint form-owned pending status without removing required durability saves.
- Viewer Events is read-only. Canonical inventory events come from published `tblInventoryEvents`; current saved box alternatives and current local held shipments are supplements, not immutable design/hold history.
- Production event labels are deferred until Production workflow review.

# Evidence and traceability

- Symptom: aggregate reference clipped after multiple BOLs -> cause: fixed-height MSForms ListBox -> requirement: complete readable reference -> protection: `Receiving.AggregateReferenceDetail` plus packaged Receiving regressions.
- Symptom: remote NAS sign-in appeared frozen -> cause: synchronous Windows SMB authentication -> requirement: progress before blocking I/O -> protection: `ServerConnection.ProgressBeforeBlockingIO`.
- Symptom: no cross-role event overview -> cause: Viewer exposed only inventory snapshot levels -> requirement: read-only Events page -> protection: focused Events checks and packaged `RunInventoryViewerEventsForTest` proof.
- Focused behavioral RED: 0/5. GREEN: 5/5.
- Packaged Inventory Viewer: one published Receipt event displayed through the real form action; ten columns, read-only, reused generation, snapshot hash unchanged.
- Packaged XLAM 74/74; live role workflows 47/47; clean ordered Release 1 chain 30/30; reviewed cleanup 11/11.
- Static baseline: 150 components, 4,729 procedures, 968 scanner candidates, 8 literal `Application.Run` targets, 47 unresolved expressions, 184 duplicate-body groups.
- Two full-chain attempts passed all business assertions but hit the known blank `/automation -Embedding` cleanup exception; after stopping only the verified test-owned process, each clean rerun passed 30/30.

# Do Not Repeat

- Do not attempt per-row ListBox height or wrapping; MSForms does not support it.
- Do not use unsupported window hooks to hide or reparent Excel's native Saving UI.
- Do not make Viewer a canonical writer or open authority workbooks for mutation.
- Do not read Box Designs only from an open Shipping operator workbook; the supplement reads the canonical Shipping BOM workbook read-only.

# Assumptions to Re-verify

- Visible remote-NAS timing depends on Windows/network conditions; automation proves pre-call feedback, not uninterrupted Excel responsiveness during the OS call.
- Current Box Designs and Held Shipments are activity snapshots, not full historical revision streams.

# Open questions and blockers

- Visible UAT remains for remote connection progress, four-reference aggregate detail, Viewer Events, and remaining native Saving-window counts.
- Full immutable history for box-design revisions and hold/release transitions would require a new Domain event contract if later requested.

# Immediate next action

Load the deployed add-ins, connect to `WHT7025AE`, verify the pre-connection status, select the four-BOL aggregate row, and open Inventory Viewer > Events to compare receipt/disposition/boxing/shipping activity with the operator forms.

# Critical references

- `src/Core/Forms/frmWarehouseConnection.frm`: `mBtnConnect_Click`
- `src/Core/Modules/modRoleEventWriter.bas`: `ConnectWarehouseStorageForCapability`
- `src/Core/Modules/modWarehouseSync.bas`: `WriteSnapshotEventRows`
- `src/Core/Modules/modInventoryViewerData.bas`: `LoadCurrentInventoryEventViewerData`
- `src/Operations/Forms/frmInventoryViewer.frm`: `ApplyViewerTab`, `RefreshEvents`
- `src/Receiving/Forms/frmReceiving.frm`: `ShowSelectedAggregateReferences`
- `src/Shipping/Modules/modTS_Shipments.bas`: `LoadShippingViewerSupplementEvents`
- `tests/tooling/Test-Plan022Slice4wOperatorResponsivenessAndEvents.ps1`
- `tools/validate_inventory_viewer.ps1`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
