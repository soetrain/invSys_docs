# Shipment Add ROW 89 repaired then lost before local apply

## Current failure

Shipments Add still fails on the Arctic-Raptor/operator test machine after Refresh shows shippable NAS inventory.

Latest user-facing error:

```text
invSys ROW 89 not found.

Completed in 2,922 ms.
00000 ms CommitCurrentLine ADD start
00000 ms CommitCurrentLine resolved selected row/operator
02922 ms CommitCurrentLine backend call
```

This is different from the prior error:

```text
Package ROW 89 not found in invSys. RepairDebug: wb=Book1;
sheet=InventoryManagement; table=invSys; rows=90; rowCol=1;
totalInvCol=14; overrideQty=9; item=T28.
```

The prior diagnostic proved:

- Latest add-in was loaded because `RepairDebug` appeared.
- Add was operating on `Book1!InventoryManagement!invSys`.
- The local `invSys` table existed with `rows=90`.
- The expected `ROW` column resolved as column `1`.
- The expected `TOTAL INV` column resolved as column `14`.
- The form/backend had enough visible NAS/display data to repair: `overrideQty=9`, `item=T28`.

## Important architecture boundary

`NAS Inv` is authoritative read-model/server truth and must not be mutated by Projected Inv, overlays, or local Shipments math.

The local `InventoryManagement!invSys` row involved here is being used only as the local Add/reservation guard for:

- floor validation,
- local `SHIPMENTS` lock/staged quantity,
- reserve event creation.

This local row repair must not become a NAS inventory writeback path.

## Recent attempted fixes

Relevant recent commits in `invSys_fork`:

```text
42dba60 Clear clipboard after workbook surface table creation
21d5d0a Directly seed missing Shipment Add row
80e7751 Use repaired Shipment Add invSys row directly
0c842a7 Add Shipment Add row repair diagnostics
fb5f694 Harden Shipment Add NAS row repair
35fe156 Repair Shipment Add from selected NAS value
9ee36b2 Repair Shipment Add from visible NAS rows
0cff469 Force Shipments Refresh to rebuild BOM view
```

Key current code path:

- `frmShipmentsTally.CommitCurrentLine`
  - captures selected Projected Inv via `SelectedShippableProjectedInventoryText`
  - captures selected NAS Inv via `SelectedShippableNasInventoryText`
  - calls `modTS_Shipments.ShipmentsFormCommitLine`
- `ShipmentsFormCommitLine`
  - writes/updates the selected line in `ShipmentsTally`
  - builds `versionAvailabilityOverrides`
  - builds `nasInventoryOverrides`
  - calls `BuildSelectedShipmentRowsDeltas`
  - then calls `ApplyShipmentDeltasLocal`
- `BuildSelectedShipmentRowsDeltas`
  - calls `RepairMissingShipmentInvSysRowFromNasOverride` when ROW is missing
  - now uses the repaired `ListRow` directly for delta creation
- `ApplyShipmentDeltasLocal`
  - re-looks up the same ROW using `FindInvListRowByRowValue`
  - current user error `invSys ROW 89 not found.` appears to come from this later step

Relevant source locations in `src/Shipping/Modules/modTS_Shipments.bas`:

- `ShipmentsFormCommitLine` around Add path
- `BuildSelectedShipmentRowsDeltas`
- `RepairMissingShipmentInvSysRowFromNasOverride`
- `FindInvListRowByRowValue`
- `ApplyShipmentDeltasLocal`

## What seems likely

The latest error suggests the delta builder got past the prior `Package ROW 89 not found in invSys` failure, but the local apply step did not find ROW 89 afterward.

That implies one of these is happening:

1. `RepairMissingShipmentInvSysRowFromNasOverride` returns a `ListRow`, but the `ROW` value is not actually persisted in the `invSys` table in a way `FindInvListRowByRowValue` can find later.
2. The repaired row is valid enough for delta creation, but `ApplyShipmentDeltasLocal` is resolving a different `invLo` instance/table than the builder used.
3. The row is being created into a blank/hidden/stale table area that is not counted by `ListRows` or is immediately changed by another hydration/reconcile call.
4. `FindInvListRowByRowValue` is still not robust against how the local row is written in the real workbook.
5. `Book1` may be the wrong workbook identity for the operational flow; the form may be running against an unsaved/generated workbook instead of the expected operator workbook.

## Clipboard side issue

The clipboard was repeatedly filling with table headers such as:

```text
GUID    USER    ACTION    ROW    ITEM_CODE    ITEM
```

and:

```text
PROCESS    DIAGRAM_ID    INPUT/OUTPUT    INGREDIENT ...
```

Commit `42dba60` applied the likely fix from guidance:

- added `ClearClipboardSurface` to `modRoleWorkbookSurfaces`
- clears both `Application.CutCopyMode` and Windows clipboard through late-bound `Forms.DataObject`
- calls it after `ListObjects.Add` in `EnsureTableSurface`
- calls it after `RebuildTableAtSurface`
- calls it after Shipping surface arrangement

There is also a Shipping-specific clipboard clear after `lo.Range.Cut Destination:=dest`, but the main suspected source was Core workbook surface table creation.

## Questions for expert review

1. In this flow, how can `BuildSelectedShipmentRowsDeltas` accept a repaired `ListRow` but `ApplyShipmentDeltasLocal` immediately fail to re-find `ROW 89` in the same local `invSys` table?

2. Should `ApplyShipmentDeltasLocal` stop re-looking up rows by `ROW` and instead accept already-resolved row references or row indexes from `BuildSelectedShipmentRowsDeltas`?

3. Is `RepairMissingShipmentInvSysRowFromNasOverride` writing the row incorrectly by using `FirstBlankListRowShipping(invLo)` / `invLo.ListRows.Add` / `WriteValue`? Should it instead resize the table and write through `invLo.ListColumns("ROW").DataBodyRange.Cells(...)` or another safer pattern?

4. Is there a chance `WriteValue lr, "ROW", 89` is silently failing in the real workbook because `InventoryManagement` is hidden/protected, filtered, malformed, or has duplicate/invalid headers?

5. The diagnostic showed `wb=Book1`. Is `Book1` the expected operator workbook for Shipments, or does this indicate `ResolveOperatorWorkbook` / `ShipmentsWorksheetForWorkbook` is operating on the wrong workbook?

6. If the local `invSys` row is only needed for Add/floor/lock, should this path stop depending on `InventoryManagement!invSys` entirely and instead validate/lock from visible NAS + `tblShippingReservations`/Shipments rows?

7. What exact additional diagnostic should be added to confirm whether `ROW 89` is present immediately after `RepairMissingShipmentInvSysRowFromNasOverride` returns and immediately before `ApplyShipmentDeltasLocal` validates deltas?

8. Should `ApplyShipmentDeltasLocal` perform its own last-chance row seed from delta metadata when `FindInvListRowByRowValue` fails, or would that hide a deeper workbook/table identity bug?

## Suggested next diagnostic if needed

Add one temporary report line after `RepairMissingShipmentInvSysRowFromNasOverride` and before `ApplyShipmentDeltasLocal`:

- workbook name/path
- table parent sheet
- `invLo.ListRows.Count`
- `FindInvRowIndexByRow(invLo, 89)`
- `FindInvListRowByRowValue(invLo, 89) Is Nothing`
- first 10 and last 10 values from `invSys[ROW]`
- whether any row has `ITEM = T28`
- whether any row has `TOTAL INV = 9`

The most useful direct proof would be: after the repair function writes the row, dump the exact row index/value it wrote and then immediately call the same finder used by `ApplyShipmentDeltasLocal`.
