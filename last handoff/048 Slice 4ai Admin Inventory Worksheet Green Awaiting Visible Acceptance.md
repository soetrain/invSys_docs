# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4ai advances D14 by adding a captured-workbook Admin inventory-list workbench
while preserving the accepted single-item editor and all packaged roles.

# Current verified state

- Last verified: 2026-08-26. Code `main` is at `59892af`; docs `main` is at
  `f489bc3` before this handoff commit.
- Slice 4ai automated gates are complete; visible operator acceptance is open.
- Excel is closed. The five rebuilt add-ins and current manifest are committed.
- Intentionally uncommitted user changes remain in code `AGENTS.md` and the
  NAS/server note at the top of docs `invSys-Design-v4.11.md`.

# Decisions and constraints

- The Ribbon control remains `btnAdminAddInventoryItem` with callback
  `modAdmin.Add_InventoryItem` and `ADMIN_MAINT`; its label is now **Add/Edit
  Inventory Items**.
- `frmAddInventoryItem` preserves single-item Add/Edit and adds independent
  **Create Inventory Table** and **Upload Selected Inventory Table** actions.
- Tables are uniquely named `invSys_Inventory_###` on `invSys Inventory
  Editor` in the saved workbook captured at launcher entry. They are staging,
  not authority; `System_Key` and `ROW` headers are prohibited.
- Upload resolves the selected table, normalizes headers, accepts extra columns
  as custom fields, preflights every pending row, then uses existing Admin
  catalog/event/processor paths. ADD codes are generated; EDIT requires exact
  Item Code and Edit Reason. Quantity modes include COUNTED, UTILITY, SERVICE,
  and NOT COUNTED.

# Evidence and traceability

- Focused RED: `1/8` passing, `7/8` behavioral RED; focused GREEN: `8/8`.
- Packaged XLAM actions `76/76`; Ribbon/compile `142/142`; live roles `47/47`;
  Release 1 full chain `30/30`.
- Launcher source/default packaged checks `24/24` and `3/3`; dedicated NAS
  runtime `16/16`; deterministic static baseline `19/19`; reviewed growth
  ratchets `13/13`.
- Slice 4ah/4ag focused regressions remain `5/5` and `6/6`.
- Static inventory: 6 packages, 154 components, 5,140 procedures, 1,040
  candidates. Full evidence is in
  `tests/integration/plan022_slice4ai_admin_inventory_worksheet_green_results.md`.

# Do Not Repeat

- Whole-sheet `Cells.EntireColumn.AutoFit` caused Excel to consume CPU and hang;
  table-range `lo.Range.Columns.AutoFit` is the bounded implementation.
- A packaged compile dialog was caused by an omitted `categoryValue` argument
  in the dry-run fixture, not the form callback; the fixture is corrected.
- Early RPC failures in live/full-chain harness attempts were rerun cleanly and
  are superseded by the GREEN results above.

# Assumptions to Re-verify

- The operator workbook remains saved and is reopened after installing the
  current packages.
- The configured warehouse UOM catalog contains the values used in visible
  table rows.

# Open questions and blockers

- Visible Slice 4ai acceptance is unresolved: create a table, enter/paste rows,
  upload the selected table, and verify the result in Inventory Viewer.
- Earlier Production visible checkpoints remain governed by Plan 022; no new
  Production behavior was changed in Slice 4ai.

# Immediate next action

Restart Excel, open the saved Admin/operator workbook, use **Add/Edit Inventory
Items** to create an Inventory table, fill or paste ADD/EDIT rows, select a cell
in that table, reopen the form to upload it, then refresh Inventory Viewer.

# Critical references

- `src/Admin/Forms/frmAddInventoryItem.frm`
- `src/Admin/Modules/modAdmin.bas`
- `src/Admin/Modules/modAdminInventoryWorksheet.bas`
- `tools/validate_phase6_packaged_xlams.ps1` check
  `Admin.InventoryWorksheetActions`
- `tests/tooling/Test-Plan022Slice4aiAdminInventoryWorksheet.ps1`
- Architecture v4.11 D14 **Admin inventory worksheet workbench**
- Plan 022 **Slice 4ai -- Admin bulk inventory worksheet staging**
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` version 1.27
