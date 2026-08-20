# Goal and release outcome

Complete Inventory Viewer event coverage by showing the Shipping **Remove** action that returns locked inventory to warehouse availability.

# Current verified state

- Code branch `main`, implementation commit `1dac3df` (`Show Shipping Remove in inventory events`).
- Documentation branch `main`, contract commit `9a7861f` (`Document Shipping Remove viewer event`).
- Active slice: Plan 022 Slice 4w follow-up — Shipping Remove event completeness.
- Focused RED/GREEN, package rebuild, packaged Viewer action, Shipping regressions, complete packaged suite, and maintenance gates are complete.
- Excel is fully closed. Last verified: 2026-08-20.
- The user's pre-existing code `AGENTS.md` and documentation `invSys-Design-v4.11.md` edits remain uncommitted and were not modified.

# Decisions and constraints

- Shipping Remove already writes canonical `SHIP_RELEASE`; no new mutation or event type was introduced.
- Viewer renders `SHIP_RELEASE` as operator-facing **Remove**, meaning the exact shipment reservation was released and locked inventory returned to warehouse availability.
- The event is read-only and comes from the published `tblInventoryEvents` snapshot projection.

# Evidence and traceability

- Symptom: Viewer used the ambiguous label `Shipment Released` -> cause: friendly-name mapping obscured the Shipping control that emitted the event -> requirement: Event Viewer includes Remove -> protection: `Viewer.Events.RemoveRelease` and the packaged `TestEventsReport` assertion.
- Focused behavioral RED: 5/6, with only `Viewer.Events.RemoveRelease` failing. GREEN: 6/6.
- Packaged Viewer action displayed two published rows including exactly one Remove row, exposed ten read-only columns, reused the same form generation, and left the snapshot byte-for-byte unchanged.
- Shipping/Boxing stabilization: 11/11; Shipping post-send: 4/4; packaged XLAM suite: 74/74; reviewed cleanup: 11/11.
- Static baseline: 150 components, 4,729 procedures, 968 scanner candidates, 8 literal `Application.Run` targets, 47 unresolved expressions, and 185 duplicate-body groups.

# Do Not Repeat

- Do not add a second Remove event or infer removal from UI state; use the existing canonical `SHIP_RELEASE` record.
- Do not change quantity authority for this display correction; release changes locked/available state without changing total inventory.

# Assumptions to Re-verify

- Visual UAT should confirm the Events row label reads **Remove** after using Shipping Remove against a held shipment.

# Open questions and blockers

- No implementation blocker remains for this correction; operator-side visual UAT is pending.

# Immediate next action

Hold a shipment line, click Shipping **Remove**, refresh Inventory Viewer > Events, and verify a read-only **Remove** row appears for the released reservation.

# Critical references

- `src/Core/Modules/modInventoryViewerData.bas`: `ViewerFriendlyEventType`
- `src/Shipping/Modules/modShippingPostingService.bas`: `EVENT_TYPE_SHIP_RELEASE`
- `src/Operations/Forms/frmInventoryViewer.frm`: `TestEventsReport`
- `tests/tooling/Test-Plan022Slice4wOperatorResponsivenessAndEvents.ps1`
- `tools/validate_inventory_viewer.ps1`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
