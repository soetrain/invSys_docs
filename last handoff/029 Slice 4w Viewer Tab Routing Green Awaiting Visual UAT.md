# Goal and release outcome

Make Inventory Viewer expose exactly Inventory and Events, with Events routed to the read-only operator event log rather than inventory levels.

# Current verified state

- Code branch `main`, implementation commit `7c24650` (`Fix inventory viewer tab routing`).
- Documentation branch `main`, contract commit `bd664a9` (`Document exact inventory viewer tabs`).
- Active slice: Plan 022 Slice 4w follow-up — Viewer tab construction and routing.
- Focused RED/GREEN, package rebuild, real packaged Viewer action, Shipping regressions, packaged suite, and maintenance gates are complete.
- Excel is fully closed. Last verified: 2026-08-20.
- The user's pre-existing code `AGENTS.md` and documentation `invSys-Design-v4.11.md` edits remain uncommitted and untouched.

# Decisions and constraints

- A dynamically added MSForms TabStrip already contains two native pages. `frmInventoryViewer` relabels those pages Inventory and Events; it must not append two more pages.
- The public Events action selects the visible Events page and reads the published `tblInventoryEvents` projection. With no published events it should display zero rows, not inventory.

# Evidence and traceability

- Symptom: visible tabs were `Tab1`, `Tab2`, `Inventory`, `Events`, and Events showed inventory -> cause: `BuildLayout` appended Inventory/Events to two automatic pages while routing by index 1 -> requirement: exactly two correctly routed tabs -> protection: `Viewer.Tabs.ExactlyInventoryAndEvents` and packaged `RunInventoryViewerEventsForTest` assertions.
- Focused behavioral RED: 6/7; only exact tab construction/routing failed. GREEN: 7/7.
- Packaged Viewer reports `TabCount=2`, `TabCaptions=Inventory,Events`, `SelectedTab=Events`, two isolated event rows, ten event columns, and unchanged snapshot hash.
- Shipping/Boxing: 11/11; Shipping post-send: 4/4; packaged XLAM suite: 74/74; reviewed cleanup: 11/11.
- Static baseline remains 150 components, 4,729 procedures, and 968 scanner candidates.

# Do Not Repeat

- Do not call `Tabs.Clear`; this MSForms collection does not support it and the packaged run opened the VBA runtime UI.
- Do not append Inventory and Events with `Tabs.Add` to a fresh runtime TabStrip; reuse its two native pages.

# Assumptions to Re-verify

- Visual UAT should show only Inventory and Events after a clean Excel restart with the rebuilt package.

# Open questions and blockers

- No automated blocker remains; clean-restart visual UAT is pending.

# Immediate next action

Restart Excel, open Inventory Viewer, verify only Inventory and Events are present, then click Events and confirm event headers and event rows (or an empty event list before new actions).

# Critical references

- `src/Operations/Forms/frmInventoryViewer.frm`: `BuildLayout`, `ApplyViewerTab`, `TestEventsReport`
- `tests/tooling/Test-Plan022Slice4wOperatorResponsivenessAndEvents.ps1`
- `tools/validate_inventory_viewer.ps1`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
