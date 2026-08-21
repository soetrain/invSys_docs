# Goal and release outcome

Stabilize Inventory Viewer maximize/minimize behavior, make event dates readable, and prove explicit Refresh replaces stale rows with the newest published events while preserving a bounded R1 projection.

# Current verified state

- Code branch `main`, source/evidence commit `ddf3567` (`Guard viewer resize and refresh events`) and deployed-package commit `b40bc0b` (`Deploy verified viewer refresh correction`).
- Documentation branch `main`, specification/contract commit `f9296d7` (`Specify bounded and future event viewers`).
- Active slice: Plan 022 Slice 4w follow-up — Viewer resize, timestamp, and refresh correction.
- Focused RED/GREEN, Core/Operations rebuild, real packaged Viewer refresh, complete packaged suite, Shipping regressions, and maintenance gates are complete.
- Excel is fully closed. Last verified: 2026-08-21.
- User grants standing permission to close Excel whenever builds/tests require it on X1-PRO-AI. Preserve work by saving an already-named unsaved workbook before closing; do not save a transient blank workbook.
- The user's pre-existing code `AGENTS.md` and the NAS/server note at the top of `invSys-Design-v4.11.md` remain outside these commits and were not altered.

# Decisions and constraints

- Operations anchoring skips minimum form-size enforcement while the native window is minimized or maximized, but retains control anchoring; residual error 384 is contained so restore can reapply layout.
- Event timestamps render as `yyyy-mm-dd hh:mm:ss` rather than Excel serial values.
- R1 Events remains a read-only bounded snapshot projection. Refresh reads a newer published snapshot but never processes or mutates authority workbooks.
- The normative specification now reserves a post-R1 comprehensive Event Viewer design covering cross-domain history, time zones, correlation, filters, pagination/retention, export, capabilities, and freshness.

# Evidence and traceability

- Symptom: run-time error 384 at `mLayout.ApplyAnchoredLayout` -> cause: minimum-size logic accessed form dimensions in a forbidden native window state -> requirement: maximize/minimize without error -> protection: `Viewer.Layout.GuardsNativeWindowState` plus packaged form loading.
- Symptom: latest receipts appeared absent -> observed rows were present below later Shipping rows, but raw serial dates obscured chronology -> requirement: readable timestamps and explicit refresh -> protection: `Viewer.Events.ReadableTimestampRefresh` and the packaged new-publication refresh proof.
- Follow-up focused RED: 7/9; only layout guarding and timestamp/refresh failed. GREEN: 9/9.
- Packaged Viewer began with two events, published a third while the form remained open, refreshed through the public Events action, displayed three rows with `BOL-VIEWER-NEW` first, and left the new snapshot byte-for-byte unchanged.
- Packaged XLAM suite: 74/74; Shipping/Boxing: 11/11; Shipping post-send: 4/4; reviewed cleanup: 11/11.
- Static baseline: 150 components, 4,729 procedures, 967 scanner candidates, 47 unresolved dynamic expressions, and 185 duplicate-body groups.

# Do Not Repeat

- Do not enforce UserForm Width/Height while its native window is minimized or maximized.
- Do not interpret the bounded R1 Events page as the future comprehensive audit viewer.
- Do not diagnose refresh solely from row position; events are newest-first and later Shipping events can precede receipts.

# Assumptions to Re-verify

- Visual maximize/minimize/restore behavior must still be confirmed in the interactive Excel session.
- Explicit Refresh should show any event already included in the latest published snapshot; unpublished/still-staged work is intentionally absent.

# Open questions and blockers

- No automated blocker remains; interactive visual UAT is pending.

# Immediate next action

Restart Excel, maximize/minimize/restore Inventory Viewer, perform a Receiving or Shipping action, click Events > Refresh, and confirm readable dates plus the newest published event without error 384.

# Critical references

- `src/Operations/ClassModules/cOperationsAnchorManager.cls`: `ApplyAnchoredLayout`
- `src/Core/Modules/modInventoryViewerData.bas`: `LoadCurrentInventoryEventViewerData`
- `src/Operations/Forms/frmInventoryViewer.frm`: `TestEventsReport`, `mBtnRefresh_Click`
- `tests/tooling/Test-Plan022Slice4wOperatorResponsivenessAndEvents.ps1`
- `tools/validate_inventory_viewer.ps1`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`: `Inventory Viewer (Release 1)`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
