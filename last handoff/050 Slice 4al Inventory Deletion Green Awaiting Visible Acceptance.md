# Slice 4al Inventory Deletion Green Awaiting Visible Acceptance

## 1. Goal and release outcome

Advance invSys Release 1 acceptance by adding audit-preserving managed-item
deletion to **Add/Edit Inventory Items** under Architecture v4.11 D14 and Plan
022 Slice 4al.

## 2. Current verified state

- Code: `main` at `444efbd`; pushed to `origin/main` on 2026-08-27.
- Docs: `main` based on `2932e8e`; Slice 4al specification, plan, controls,
  and this handoff are pending the accompanying docs commit.
- Active slice: 4al -- Add/Edit managed-inventory deletion.
- Automated gates are GREEN; visible operator acceptance remains open.
- Excel was closed after validation and before handoff.
- User-owned `AGENTS.md` and the normative-spec NAS/server note remain
  uncommitted and must be preserved.

## 3. Decisions and constraints

- **Delete Item** appears only in Edit mode and is enabled only after an exact
  catalog-backed SKU selection.
- Deletion requires confirmation and a nonblank reason.
- The action retires every active exact `System_Key` for the SKU through one
  `ADMIN_INVENTORY_ADJUST` event. Counted balances go to zero; Utility/Service
  entities receive zero-delta retirement evidence.
- `InventoryState=RETIRED` and `CATALOG_STATE=RETIRED` suppress the item from
  active projections. Keys, catalog definitions, and event history are never
  physically erased or reused.
- No worksheet bulk-delete verb was added in Release 1.

## 4. Evidence and traceability

- Focused D13 RED: `1/7`; focused GREEN: `7/7`.
- Packaged XLAM actions: `79/79`, including the real Delete handler and isolated
  counted plus Utility Domain retirement.
- Launcher source contracts: `24/24`; packaged launcher reuse: `3/3`.
- Ribbon/compile: `142/142`; live roles: `47/47`; Release 1 chain: `30/30`;
  dedicated NAS: `16/16`; deterministic static: `19/19`; reviewed cleanup:
  `13/13`.
- Protecting test: `tests/tooling/Test-Plan022Slice4alInventoryDeletion.ps1`.

## 5. Do Not Repeat

- Do not implement deletion by removing entity, catalog, or event rows.
- Do not treat a zero-delta Utility retirement as an invalid quantity.
- Generate the deterministic baseline with timestamp
  `2026-08-16T20:00:00Z`; a current timestamp fails byte comparison.

## 6. Assumptions to Re-verify

- Re-verify deployed visible behavior against a disposable item before deleting
  business inventory.
- Re-verify Events wording/details after the retirement is published and the
  Viewer is refreshed.

## 7. Open questions and blockers

- Visible acceptance is unresolved: the user must confirm the selected item
  disappears from managed inventory, Viewer, Production picker, and ordinary
  Edit search while its Inventory Adjustment evidence remains visible.

## 8. Immediate next action

Use **Add/Edit Inventory Items** > **Edit Item**, select a disposable managed
item, click **Delete Item**, confirm and enter a reason, then Refresh the form,
Viewer, and Production picker to record visible acceptance.

## 9. Critical references

- `src/Admin/Forms/frmAddInventoryItem.frm`: `mBtnDeleteItem_Click`.
- `src/Admin/Modules/modAdmin.bas`: `RetireInventoryItemForWarehouse`.
- `src/InventoryDomain/Modules/modInventoryApply.bas`:
  `ValidateRetirementPayloadLineApply`, `RebuildInventoryProjections`.
- `src/Core/Modules/modOperatorReadModel.bas`: `ApplySnapshotToInvSys`.
- `src/Core/Modules/modInventoryViewerData.bas`: `ViewerFriendlyEventType`.
- Architecture v4.11 D14, Plan 022 Slice 4al, controls v1 `btnDeleteItem`.
