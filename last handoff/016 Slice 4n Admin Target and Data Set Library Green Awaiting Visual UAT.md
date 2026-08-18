# Slice 4n Admin Target and Data Set Library Green Awaiting Visual UAT

## 1. Goal and release outcome

Give Admin the same live warehouse selector as Operations and make uploaded
demo inventory definitions persistent, selectable, and independently deletable
without conflating definition deletion with inventory depletion. This advances
v4.11 D-NAS, D14, and Plan 022 Slice 4n.

## 2. Current verified state

- Last verified: 2026-08-17.
- Code: `main`, commit `a226b25` (`Add Admin warehouse and demo dataset selectors`).
- Docs contract/plan/catalog: `main`, commit `c55cb77`
  (`Document Admin target and dataset management`).
- Automated gates are complete; visible dedicated-NAS UAT remains.
- Excel is closed. Rebuilt Core and Admin packages and the manifest are in
  `deploy/current`.
- User-owned unstaged changes remain in code `AGENTS.md` and the NAS/server
  note at the top of docs `invSys-Design-v4.11.md`; neither was committed.

## 3. Decisions and constraints

- Admin `Send To` uses the same Core target list, selected index, and public
  `RibbonWarehouseOnAction` callback as Operations.
- Uploaded definitions are validated before copying to the selected warehouse's
  `admin\demo-inventory-data-sets` library. Same-named files are rejected
  rather than overwritten.
- Uploading/selecting a definition does not add inventory. Seed applies the
  selected definition.
- Delete Data Set deletes only the selected uploaded CSV. It does not change
  previously seeded inventory. The R1 kit is immutable and its delete button is
  disabled while selected; the service also rejects attempted deletion.
- Delete Demo Inventory remains separate: it posts exact-`System_Key`
  depletion adjustments for active `DEMO-` stock and retains canonical history.

## 4. Evidence and traceability

- Admin selector focused RED: 9/10; selector/invalidation absent.
- Admin selector GREEN: 10/10; packaged Ribbon 142/142.
- Dataset library focused RED: 9/12; persistence, definition delete, and R1
  protection absent.
- Dataset library GREEN: 12/12.
- Packaged public callback: PASS for import/list, seed from stored definition,
  repeated seed, definition deletion with inventory retained, R1 protection,
  invalid-file rejection, inventory depletion, and canceled deletion.
- Packaged XLAM: 74/74.
- Ordered Receiving -> Production -> Box Maker -> Shipping -> restart and
  reconciliation: 30/30.
- Deterministic static evidence: 19/19; 150 components, 4,670 procedures, 962
  scanner candidates, 964 reviewed candidates, and 23 oversized-module
  ratchets.
- A first dataset-package run stalled because `runtimeRoot` was declared both
  as a ByRef parameter and a local variable. The duplicate was removed before
  GREEN; do not reintroduce it.

## 5. Do Not Repeat

- Do not use Delete Data Set to reverse inventory; it owns only a stored CSV.
- Do not physically delete canonical inventory/event history.
- Do not allow a supplied path outside the selected warehouse data-set library
  to reach the delete service.
- Do not overwrite an existing uploaded definition implicitly.
- Do not build packages while Excel is open.

## 6. Assumptions to Re-verify

- The account-scoped add-in registration still points to `deploy/current`.
- The selected NAS warehouse permits Admin to create and delete files beneath
  its `admin\demo-inventory-data-sets` folder.
- Both ribbon selectors visibly update immediately at the user's display scale.

## 7. Open questions and blockers

- No automated blocker remains.
- The Admin ribbon layout, cross-ribbon target refresh, and managed-library
  workflow require visible acceptance against `WHT7025AE`.

## 8. Immediate next action

Open Excel cleanly, select `WHT7025AE` from Admin `Send To`, confirm
Operations updates to the same target, then upload/select/delete a disposable
CSV definition and verify its seeded inventory remains after definition delete.

## 9. Critical references

- `tools/build-xlam.ps1`, Admin Ribbon configuration
- `src/Core/Modules/modRibbonRuntimeStatus.bas`
- `src/Admin/Forms/frmSeedInventory.frm`
- `src/Admin/Modules/modAdmin.bas`, `Seed_DemoInventory`
- `src/Admin/Modules/modAdminInventorySeed.bas`
- `tests/integration/admin_warehouse_selector_red_results.md`
- `tests/integration/demo_dataset_library_red_results.md`
- `tests/integration/admin_seed_callback_green_results.md`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, version 1.12
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4n
