# Goal and release outcome

Remember the operator's last valid Viewer Event range per Windows user so reopening the Viewer does not revert to the first-use default.

# Current verified state

- Code branch `main`, commit `ab77581` (`Remember Viewer event date range`), pushed to GitHub.
- Documentation branch `main`, contract commit `7031f89` (`Specify remembered Viewer event range`), pushed before this handoff.
- Active slice: Plan 022 Slice 4w follow-up - remembered Viewer Event range; automated gates are complete.
- Core and Operations were rebuilt and deployed to `deploy/current`.
- Excel is fully closed. Last verified: 2026-08-21.
- The user's pre-existing code `AGENTS.md` change and NAS/server note at the top of `invSys-Design-v4.11.md` remain uncommitted and preserved.

# Decisions and constraints

- `All` is only the first-use/fallback range. Each valid applied `All`, `Day`, `Week`, `Month`, or custom 1-36500 value is stored with VBA `SaveSetting` under the per-Windows-user invSys Operations preferences.
- A new Viewer form reads the preference with `GetSetting`; this storage survives Excel restart. Invalid persisted text falls back to `All`.
- The preference is local UI state. It never enters Config, Inventory, Snapshot, or another warehouse workbook.
- A range is considered applied when Viewer rendering/Refresh accepts it; invalid input is not persisted.

# Evidence and traceability

- Focused RED: 10/11; only `Viewer.Events.RemembersDateFilter` failed. GREEN: 11/11.
- Packaged Viewer applied custom 14 days, unloaded the Viewer, created generation 2, restored 14 days with three rows, and proved invalid persisted text falls back to All.
- The packaged test preserved and restored the machine's pre-test preference and left the snapshot hash unchanged.
- Packaged XLAM suite: 74/74; reviewed cleanup: 11/11.
- Static baseline: 150 components, 4,729 procedures, 965 candidates, 47 unresolved dynamic calls, and 185 duplicate-body groups.

# Do Not Repeat

- Do not store this preference in warehouse workbooks or key it to inventory identity.
- Do not persist invalid or partially typed custom values.
- Do not restore `All` unconditionally in `BuildLayout`; it is now the first-use/corrupt-value fallback only.

# Assumptions to Re-verify

- Visual UAT must confirm the user's selected range remains visible after closing/reopening Viewer and after restarting Excel.

# Open questions and blockers

- No automated blocker remains. Interactive visual UAT is pending.

# Immediate next action

Open Viewer > Events, apply a non-default range, close/reopen Viewer and then restart Excel to confirm the same range returns.

# Critical references

- `src/Operations/Forms/frmInventoryViewer.frm`: `BuildLayout`, `RenderRows`
- `tests/tooling/Test-Plan022Slice4wOperatorResponsivenessAndEvents.ps1`
- `tools/validate_inventory_viewer.ps1`
- `tests/integration/inventory_viewer_results.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`: `Inventory Viewer (Release 1)`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`: Slice 4w
