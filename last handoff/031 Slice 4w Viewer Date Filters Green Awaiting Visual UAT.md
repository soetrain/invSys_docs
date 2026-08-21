# Goal and release outcome

Add bounded operator date filters to the read-only R1 Viewer Events page without expanding it into the future comprehensive Event Viewer or adding authority writes.

# Current verified state

- Code branch `main`, commit `c766453` (`Add Viewer event date filters`), pushed to GitHub.
- Documentation branch `main`, contract commit `15bbaa2` (`Specify Viewer event date filters`), pushed before this handoff.
- Active slice: Plan 022 Slice 4w follow-up - Viewer rolling date filters; automated gates are complete.
- Core and Operations were rebuilt and deployed to `deploy/current`.
- Excel is fully closed. Last verified: 2026-08-21.
- The user's pre-existing code `AGENTS.md` change and NAS/server note at the top of `invSys-Design-v4.11.md` remain uncommitted and preserved.

# Decisions and constraints

- Events defaults to `All`; operators may choose `Day`, `Week`, `Month`, or type a positive whole number from 1 to 36500 and select Refresh.
- Day, Week, and Month mean rolling 1-, 7-, and 30-day windows. The date range combines with Search.
- The Event range controls are hidden on Inventory, which remains text-filtered only.
- Filtering applies to the loaded read-only projection. It never processes or writes an authority workbook.
- The post-R1 comprehensive Event Viewer remains separate design work, including time-zone, retention, pagination, export, and broader cross-domain requirements.

# Evidence and traceability

- Focused RED: 9/10; only `Viewer.Events.RollingDateFilters` failed.
- Focused GREEN: 10/10 through the same public Refresh action used by the operator.
- Packaged Viewer: initial four events, refreshed five-event publication, Day=2, Week=3, Month=4, custom 14-day=3, All=5; one form generation; readable dates; Events read-only; snapshot hash unchanged.
- Packaged XLAM suite: 74/74.
- Reviewed cleanup: 11/11.
- Static baseline: 150 components, 4,729 procedures, 966 candidates, 47 unresolved dynamic calls, and 185 duplicate-body groups.

# Do Not Repeat

- Do not remove the default `All` option; it preserves the existing complete published view.
- Do not add an Event-range handler solely for convenience if it increases procedure bloat; the existing Refresh action is the application boundary and is test-covered.
- Do not treat the bounded R1 page as canonical history or the future comprehensive audit viewer.

# Assumptions to Re-verify

- The new second control row and anchored list remain visually readable through normal, maximized, minimized, and restored window states.
- Existing R1 timestamps retain their current non-time-zone-aware display semantics; time-zone-aware filtering belongs to the future comprehensive Viewer design.

# Open questions and blockers

- No automated blocker remains. Interactive visual UAT is pending.

# Immediate next action

Restart Excel, open Inventory Viewer > Events, visually confirm All/Day/Week/Month/custom-number filtering plus Search through resize and maximize/minimize/restore.

# Critical references

- `src/Operations/Forms/frmInventoryViewer.frm`: `BuildLayout`, `ApplyViewerTab`, `RenderRows`, `TestEventsReport`
- `src/Operations/Modules/modInventoryViewer.bas`: `RunInventoryViewerEventsForTest`
- `tests/tooling/Test-Plan022Slice4wOperatorResponsivenessAndEvents.ps1`
- `tools/validate_inventory_viewer.ps1`
- `tests/integration/inventory_viewer_results.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`: `Inventory Viewer (Release 1)`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`: Slice 4w
