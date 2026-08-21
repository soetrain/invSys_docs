# Goal and release outcome

Keep the R1 Viewer Events page limited to meaningful operator control actions so an ordinary Shipping Add does not appear as a zero-quantity Shipment Held action.

# Current verified state

- Code branch `main`, commit `be19f04` (`Hide internal shipping reservations from events`), pushed to GitHub.
- Documentation branch `main`, contract commit `12a1152` (`Define viewer as operator action log`), pushed before this handoff.
- Active slice: Plan 022 Slice 4w follow-up - operator-action Event visibility; automated gates are complete.
- Core and Operations were rebuilt and deployed to `deploy/current`.
- Excel is fully closed. Last verified: 2026-08-21.
- The user's pre-existing code `AGENTS.md` change and NAS/server note at the top of `invSys-Design-v4.11.md` remain uncommitted and preserved.

# Decisions and constraints

- The bounded R1 Events page reports meaningful operator control actions, not backend processor or staging mechanics invoked while carrying out those actions.
- `SHIP_RESERVE` written by ordinary Shipping Add is internal reservation machinery and is excluded from the operator-facing projection.
- Actual current held-shipment supplements remain eligible to appear as `Shipment Held` after an operator Hold action.
- `SHIP_RELEASE` remains visible as `Remove` even with quantity zero because it records an explicit operator release of locked inventory.

# Evidence and traceability

- Focused RED: 11/12; only `Viewer.Events.ExcludesInternalReservation` failed. GREEN: 12/12.
- Packaged public Viewer action passed with a zero-delta `SHIP_RESERVE` fixture present but hidden, `Remove` visible, all date ranges intact, remembered custom 14 days intact, and the snapshot hash unchanged.
- Packaged XLAM suite: 74/74; reviewed cleanup: 11/11.
- Static baseline: 150 components, 4,729 procedures, 965 candidates, 8 literal `Application.Run` targets, 47 unresolved dynamic calls, and 185 duplicate-body groups.

# Do Not Repeat

- Do not map canonical `SHIP_RESERVE` to `Shipment Held`; Shipping Add is not a Hold action.
- Do not blanket-filter every zero-quantity row because `Remove` is a meaningful operator state-release action.
- Do not remove the separate `SHIP_HELD` supplement used for actual current holds.

# Assumptions to Re-verify

- Visual UAT must confirm a normal Add/Shipments Sent flow shows Shipped without a zero-quantity Shipment Held row.
- Visual UAT should confirm an actual Hold still appears as Shipment Held and Remove still appears after releasing locked inventory.

# Open questions and blockers

- No automated blocker remains. Interactive visual UAT is pending.

# Immediate next action

Run one normal Shipping Add plus Shipments Sent and one explicit Hold/Remove sequence, then refresh Viewer > Events and verify only the corresponding operator actions appear.

# Critical references

- `src/Core/Modules/modInventoryViewerData.bas`: `ViewerFriendlyEventType`
- `src/Operations/Forms/frmInventoryViewer.frm`: `TestEventsReport`
- `src/Shipping/Modules/modTS_Shipments.bas`: `LoadShippingViewerSupplementEvents`
- `tests/tooling/Test-Plan022Slice4wOperatorResponsivenessAndEvents.ps1`
- `tools/validate_inventory_viewer.ps1`
- `tests/integration/inventory_viewer_results.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`: `Inventory Viewer (Release 1)`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`: Slice 4w
