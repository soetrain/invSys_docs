# Slice 4n Demo Inventory Lifecycle Green Awaiting Visual UAT

## 1. Goal and release outcome

Provide an administrator-selected, repeat-safe demo inventory lifecycle and a
readable Receiving choice projection so a complete Release 1 operator workflow
can be tested. This advances v4.11 D14 and Plan 022 Slice 4n acceptance.

## 2. Current verified state

- Last verified: 2026-08-17.
- Code: `main`, commit `375f9fa` (`Add selectable demo inventory lifecycle`).
- Docs contract/plan/catalog: `main`, commit `4f833f4` (`Define selectable demo
  inventory lifecycle`).
- Active slice: Plan 022 Slice 4n; automated gates are complete and visible
  dedicated-NAS UAT remains.
- Excel is closed. The rebuilt Core, Operations, and Admin packages and current
  five-package manifest are in `deploy/current`.
- User-owned unstaged changes remain in code `AGENTS.md` and the NAS/server
  note at the top of docs `invSys-Design-v4.11.md`; neither was committed.

## 3. Decisions and constraints

- The Demo Inventory form selects either the built-in 24-item R1 workflow kit
  or an uploaded CSV. Upload chooses the file; Seed performs the mutation.
- A repeated seed skips an already active item-code/location/condition group.
  A missing or fully depleted group receives a new immutable `System_Key`.
- Delete requires confirmation and posts exact-key negative adjustments for
  active `DEMO-` entities. Canonical entity and event history is retained.
- CSV requires `ITEM_CODE`, `ITEM`, `QTY`, `UOM`, and `LOCATION`; optional
  fields are `CONDITION`, `DESCRIPTION`, `CATEGORY`, and `VENDOR`. Validation
  is all-or-nothing; item codes must begin `DEMO-` and quantities are positive.
- Receiving search aggregates by item code/UOM/location/condition and hides
  nonpositive totals. Its retained representative key is only for catalog
  lookup; a confirmed receipt creates its own durable key.

## 4. Evidence and traceability

- Symptom: repeated demo seeds appeared as duplicate-looking rows in Receiving.
- Root cause: each prior seed intentionally generated new durable keys and the
  Receiving picker projected each entity separately.
- Focused RED records four absent lifecycle contracts, repeated-seed growth,
  stale zero-quantity projection rows, and missing selected-data-set routing in
  `tests/integration/demo_inventory_lifecycle_red_results.md`.
- Source lifecycle contract: 9/9.
- Receiving form-action aggregation: 2/2.
- Generate Warehouse/Create Warehouse lifecycle: 15/15.
- Packaged public Demo Inventory callback: PASS for built-in seed, repeat seed,
  exact-key delete, selected CSV seed, repeat CSV, invalid CSV rejection, and
  canceled delete.
- Packaged XLAM: 74/74; packaged Ribbon: 140/140.
- Ordered Receiving -> Production -> Box Maker -> Shipping -> restart and
  reconciliation: 30/30.
- Deterministic static evidence: 19/19; 150 components, 4,662 procedures, 959
  scanner candidates, 961 reviewed candidates, and 23 oversized-module
  ratchets.
- Broader Phase 6 ranges still contain pre-existing unrelated failures at
  Shipping surface tests 217/218 and Inventory picker test 280; they were not
  introduced or changed by Slice 4n.

## 5. Do Not Repeat

- Do not physically delete inventory or event rows to implement Demo Delete.
- Do not identify demo inventory only by display fields when posting a
  depletion; use every exact active `System_Key`.
- Do not mutate inventory when the operator merely chooses an upload file.
- Do not run package builds while Excel has invSys add-ins or operator
  workbooks open.

## 6. Assumptions to Re-verify

- Account-scoped add-in registration still points to `deploy/current`.
- The uploaded CSV picker is readable at the operator's Windows display scale.
- The dedicated NAS test warehouse may contain prior depleted demo history;
  this is expected and must not be mistaken for active duplicate stock.

## 7. Open questions and blockers

- No automated blocker remains.
- The revised runtime-generated form and the operator-selected CSV path require
  visible acceptance against `WHT7025AE`.

## 8. Immediate next action

Open Excel cleanly, select `WHT7025AE`, open Admin **Demo Inventory**, verify
the built-in/CSV data-set selector, seed the built-in kit twice, and confirm the
second result creates no active duplicate groups before beginning full UAT.

## 9. Critical references

- `src/Admin/Forms/frmSeedInventory.frm`
- `src/Admin/Modules/modAdmin.bas`, `Seed_DemoInventory`
- `src/Admin/Modules/modAdminInventorySeed.bas`
- `src/Receiving/Modules/modTS_Received.bas`,
  `LoadReceivingFormInventoryForWorkbook`
- `src/Core/Modules/modInventoryViewerData.bas`
- `tests/integration/admin_seed_callback_green_results.md`
- `tests/integration/demo_inventory_lifecycle_red_results.md`
- `tests/integration/slice14_results.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`, D14
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, version 1.11
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4n
