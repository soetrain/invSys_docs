# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4ah makes an Admin inventory combo choice bind the exact catalog SKU so the
operator can edit Filtered Water to Utility without a false missing-selection
error, then resume Production saved-workbook acceptance.

# Current verified state

- Last verified 2026-08-26: code `main` at `da874c9`; plan/control docs at
  `8698e35`.
- Active slice: Plan 022 Slice 4ah, automated GREEN; visible Admin retest is
  pending. Production Slice 4ag visible Actual Output retest also remains open.
- Rebuilt five-package set is in `deploy/current`; Excel is closed.
- Preserved uncommitted user work: code `AGENTS.md` and the NAS/server note at
  the top of Architecture v4.11.

# Decisions and constraints

- A visible combo value is not item identity. Choosing a combo row now commits
  its exact catalog SKU before typed-search filtering can rebuild the list.
- Combo-dropdown and search-results choices share the same SKU-backed field
  loader. Typed text alone continues to filter and does not invent a selection.
- Qty mode **Utility** emits `TRACK_QTY=FALSE` and `ITEM_KIND=UTILITY`, sets no
  numeric target quantity, and remains attached to the selected SKU.
- This fixes the existing Admin/control contract; it is not an Architecture
  v4.11 change. Catalog and Inventory Domain authority remain unchanged.

# Evidence and traceability

- Focused D13 RED `0/5`; focused GREEN `5/5`.
- Packaged Admin form test invokes the real combo Change handler and records
  `ComboSelected=True`, `FieldsLoaded=True`, `UtilityReady=True`, and
  `ValidationReady=True`.
- Regressions: XLAM `75/75`, Ribbon/compile `142/142`, live roles `47/47`,
  ordered Release 1 chain `30/30`, launcher contracts `24/24`, dedicated NAS
  `16/16`, deterministic static `19/19`, reviewed growth `13/13`.
- Static metrics: 153 components, 5,096 procedures, 1,038 candidates; dynamic
  call ratchets did not regress.

# Do Not Repeat

- Do not infer that combo display text means `mSelectedEditSku` is bound.
- Do not let a real dropdown selection fall through the typed-search path that
  resets `ListIndex`.
- Do not create a second catalog or inventory authority to support Utility.

# Assumptions to Re-verify

- Confirm normal Excel startup loads the `deploy/current` Admin hash rather
  than an already-open cached add-in.
- The selected Filtered Water catalog row currently carries category Utility;
  visible Save/refresh behavior is not accepted until the operator retests it.

# Open questions and blockers

- Visible confirmation that Filtered Water fields load and Save Item accepts
  Utility without **Choose an inventory item to edit** is unresolved.
- Production Slice 4ag visible confirmation of Actual Output/Last Actual and
  the broader Slice 4 saved-workbook/NAS acceptance remain unresolved.

# Immediate next action

Open Admin Add/Edit Inventory Item, choose Filtered Water from the Inventory
item dropdown, confirm Item name populates, select Utility, save with a reason,
then refresh Viewer and the Production picker before returning to Production
Run - List acceptance.

# Critical references

- `src/Admin/Forms/frmAddInventoryItem.frm:mCmbEditItem_Change`
- `src/Admin/Forms/frmAddInventoryItem.frm:CommitSelectedEditItemFromCombo`
- `src/Admin/Forms/frmAddInventoryItem.frm:CommitEditItemSelection`
- `src/Admin/Forms/frmAddInventoryItem.frm:TestEditItemComboSelectionContract`
- `src/Admin/Modules/modAdmin.bas:InventoryEditSelectionContractForAutomation`
- `tests/tooling/Test-Plan022Slice4ahAdminEditSelection.ps1`
- `tests/integration/plan022_slice4ah_admin_edit_selection_green_results.md`
- `tools/validate_phase6_packaged_xlams.ps1:Admin.EditItemComboSelection`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md` Slice 4ah
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` Slice 4ah and Admin controls
