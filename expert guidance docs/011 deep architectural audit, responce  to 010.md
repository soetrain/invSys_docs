Now I have all the information needed for a thorough expert response. Let me write the complete guidance document.

***

This is a deep architectural audit. Based on reading the full source — `frmShipmentsTally.frm`, `modTS_Shipments.bas` (18,684 lines, 744 KB), and the guidance doc — here are precise answers to all seven questions in `010 Shipments unusable performance and NAS inventory corruption.md` , grounded in the actual code .

***

## Q1 — The 65-Second Load Call Chain

**`InitializeFromShipping`** is the entry point . The slow path flows as follows:

1. `LoadShippables(operatorWb)` → `ShipmentsFormLoadShippables` → **`BoxMakerFormLoadSavedBoxes(operatorWb)`** — scans the `ShippingBOM` runtime table. If the local `ShippingBOMView` is missing or stale, the fallback in the form calls `EnsureShippingBomViewPopulated(wb, bomReport, True)`, which triggers **`RefreshShippingBomViewForWorkbook`** — a full BOM view rebuild.

2. `BoxMakerFormLoadShippableVersionInventory(savedBoxes, operatorWb)` — for every saved box, calls **`BoxMakerShippingBomSourceTable`**, which calls **`OpenShippingBomWorkbook`** → **`OpenWorkbookHiddenShipping`** on the NAS runtime path (`warehouseId.invSys.Data.Inventory.xlsb` or the ShippingBOM workbook) . **This is opening a hidden NAS workbook on every shippable load.** It also calls **`BoxMakerFormLoadBoxVersionInventory`** per item, which reads version inventory from the BOM log tables — one round-trip per box row.

3. `ResolveCurrentInventoryValue` is called per box item and walks through: `invSys` table → `BuildRuntimeSnapshotInventoryCache` (which may open the NAS snapshot workbook) → `invSysData_Shipping` table → `Check_invSys` table. Each fallback adds scan cost.

4. `LoadShipmentState` runs `ShipmentsFormLoadLines` twice (active and hold), each scanning the `ShipmentsTally` sheet's `TABLE_SHIPMENTS` / `TABLE_NOTSHIPPED` list objects.

5. `EvictOrphanedActiveOverlays` iterates all shippable rows × all shipment list rows.

6. `RefreshProjectedShippableInventory` calls `ShipmentsFormLoadNasReservationTotals()` and then iterates all shippable rows again.

7. `EnforceShippingSupportSheetsHidden` is called twice (start + cleanup) and accesses worksheets by name, which can trigger Excel recalculation events on busy workbooks.

**Root cause of 65 seconds:** `BoxMakerShippingBomSourceTable` can open a transient NAS workbook (`openedTransient = True`). On a UNC/NAS path this cold-open can take 10–40 seconds alone. Multiply by the fallback retry when `ShippableRowsLoaded` returns False (a second full call chain) and the cost compounds.

***

## Q2 — Paths That Write `TOTAL INV` or Derive Visible NAS Qty from Non-NAS Sources

There are **four active violation paths** in the current code:

**Path A — `CommitCurrentLine` ADD (the "13→0" culprit)**

In `frmShipmentsTally.CommitCurrentLine`, for ADD/UPDATE: 

```vba
Set ws = modTS_Shipments.GetShipmentsTallyWorksheet(operatorWb)
If Not ws Is Nothing Then modTS_Shipments.ShipmentsFormHydrateInvSysFromShippables ws, mShippables
```

This runs **before** `ShipmentsFormCommitLine`. It calls `ShipmentsFormHydrateInvSysTableFromShippables`, which writes `invLo("TOTAL INV")` from `mShippables(r, 8)` — the **projected quantity** (NAS minus active shipment qty), not the NAS quantity. So when you add an item, the projected display quantity (e.g. `13 - 13 active = 0`) gets written to `TOTAL INV` in `InventoryManagement!invSys` **before** the reserve event fires. This is why 13 drops to 0.

Inside `ShipmentsFormCommitLine`, the new-reserve (non-delta) path also calls `GetWritableShippingInvSysTable`, which chains to `ReconcileShippableTotalsFromVersionInventory` (see Path D below).

**Path B — `To Shipments` / `Shipments Sent` via `GetWritableShippingInvSysTable`**

`ShipmentsFormRunToShipmentsRows` and `ShipmentsFormRunShipmentsSentRows` both call `GetWritableShippingInvSysTable(ws, report)` at the top of their execution . This always runs `ReconcileShipmentStagingFromShipmentLines` and `ReconcileShippableTotalsFromVersionInventory` on `invSys` — every single time a Send or Stage action fires.

**Path C — `ShipmentsFormHydrateInvSysTableFromShippables` writes TOTAL INV from version logs**

When `existingTotal` is blank or zero but `totalInvQty > 0`, it writes the version-log total to `TOTAL INV` :

```vba
colTotalInv = ColumnIndex(invLo, "TOTAL INV")
If colTotalInv > 0 Then invLo.ListRows(existingIdx).Range.Cells(1, colTotalInv).Value = totalInvQty
```

`totalInvQty` is sourced from `mShippables(r, 8)` (projected) with fallback to `mShippables(r, 4)` (NAS text), then further fallback to `BoxMakerFormLoadBoxVersionInventory` (log scan). If the NAS column is blank (read-model not yet refreshed), the log scan result gets written to `TOTAL INV` — this is the "19→24" unrelated item corruption: another item's version log had a higher count than the stale NAS value and overwrote it.

**Path D — `ReconcileShippableTotalsFromVersionInventory` overwrites TOTAL INV unconditionally**

This function calls `BoxMakerFormLoadSavedBoxes()` (no workbook argument — uses the XLAM-scoped active context) and `BoxMakerFormLoadBoxVersionInventory` per box row, summing all version quantities. Then: 

```vba
If availableVersionQty > NzDbl(invLo.DataBodyRange.Cells(invIdx, cTotal).Value) + 0.0000001 Then
    invLo.DataBodyRange.Cells(invIdx, cTotal).Value = availableVersionQty
End If
```

If version-log total > current `TOTAL INV`, it **inflates TOTAL INV to the log-derived value.** This is the direct invariant violation — version inventory logs are used to repair the authoritative NAS quantity upward.

***

## Q3 — Should `GetWritableShippingInvSysTable` Be Split?

**Yes — remove it from all display and form-open paths immediately.**

The function has four responsibilities bundled together:
1. Locate or create `InventoryManagement!invSys` (sheet creation side effect)
2. Optionally call `modRoleWorkbookSurfaces.EnsureInventoryManagementSurface` (triggers sheet rebuild/activation)
3. Hydrate `invSys` from shipment lines or read-model if empty
4. Reconcile staging and totals from version logs

The proposed split is correct and should be implemented as:

```vba
' Read path — no mutation, no EnsureInventoryManagementSurface
Private Function GetNasReadModelForDisplay(ByVal wsShip As Worksheet) As ListObject
    ' Returns invSys or invSysData_Shipping as read-only; never hydrates, never reconciles

' Local floor-guard path — only reads SHIPMENTS column, never TOTAL INV
Private Function GetLocalShipmentLocksForFloorGuard(ByVal wsShip As Worksheet) As ListObject

' Mutable local staging — allowed only from To Shipments and Shipments Sent actions
Private Function GetMutableLocalShipmentStagingTable(ByVal wsShip As Worksheet, ByRef report As String) As ListObject
    ' Calls EnsureMissingInvSysRowsFromShipmentLines and ReconcileShipmentStagingFromShipmentLines
    ' Never calls ReconcileShippableTotalsFromVersionInventory
```

The current `GetWritableShippingInvSysTable` call sites are: form load (via `BoxMakerFormLoadShippableVersionInventory` indirectly), `CommitCurrentLine` ADD, `ShipmentsFormRunToShipmentsRows`, `ShipmentsFormRunShipmentsSentRows`, and `ShipmentsFormRunShipmentsSent` (old single-selection path). All of these should be audited individually — display paths get the read variant, staging gets the mutable variant, and NAS display is sourced only from `invSysData_Shipping` or the snapshot cache.

***

## Q4 — Is `ReconcileShippableTotalsFromVersionInventory` Safe in Any Shipments Form Path?

**No — it is not safe in any Shipments form path under v4.10 and should be quarantined.**

The function's logic is: "If the sum of version-log inventory is higher than `TOTAL INV`, raise `TOTAL INV` to match." Under v4.10's invariant this is a direct violation: it means a locally-derived version-log total can override the authoritative NAS value upward whenever `TOTAL INV` is low or stale. 

There is one theoretical safe use — in BoxMaker's manufacturing commit after `BtnBoxesMade` posts new finished-good units to `TOTAL INV`. There, inflating to version-log total may be intentional because the version log is the source of truth for what was just made. But that path already has its own direct `TOTAL INV` write (line 211 message: "added X shippable units to TOTAL INV") and does not need the reconcile function to run again.

**Recommended disposition:** Remove all calls from Shipments form paths. Keep a single call from the BoxMaker `BtnBoxesMade` / `BtnBoxCreated` post-manufacturing path if needed, behind a flag or renamed function. Add a guard at the top of the function:

```vba
Private Sub ReconcileShippableTotalsFromVersionInventory(ByVal invLo As ListObject, _
                                                         Optional ByVal callerContext As String = "")
    If StrComp(callerContext, "MANUFACTURING", vbTextCompare) <> 0 Then
        Debug.Assert False ' Should never be called from Shipments form paths
        Exit Sub
    End If
```

***

## Q5 — Safest Minimal Architecture for Sub-1000ms Form Load + Add

The key insight from the code is that **the expensive work is workbook I/O** — `BoxMakerShippingBomSourceTable` opens a NAS file, `BoxMakerFormLoadBoxVersionInventory` scans BOM log tables per item, and `BuildRuntimeSnapshotInventoryCache` may open a snapshot workbook. None of this belongs in form open.

**Form load target architecture:**

```
InitializeFromShipping:
  1. Resolve operator workbook (in-memory, < 1 ms)
  2. Read ShippingBOMView from operator workbook (local table scan, < 50 ms)
  3. Read invSysData_Shipping or snapshot cache for NAS Inv column only (local, < 20 ms)
  4. Zip into mShippables array: [ROW, Box, Version, NAS Inv, UOM, Location]
  5. Read ShipmentsTally list objects for active/hold rows (local, < 20 ms)
  6. Compute Projected Inv = NAS Inv - active qty (in-memory, < 5 ms)
  7. Render listboxes
```

**What to eliminate or defer:**
- `EnsureShippingBomViewPopulated` with `forceRefresh:=True` — only run on explicit Refresh button
- `BoxMakerShippingBomSourceTable` (NAS workbook open) — blocked from form load; call only in background or on Refresh
- `ReconcileShippableTotalsFromVersionInventory` — removed from all form paths
- `GetWritableShippingInvSysTable` → `EnsureInventoryManagementSurface` — blocked from form load
- `BoxMakerFormLoadBoxVersionInventory` per-item log scan — blocked from form load

**`Add` target architecture:**

```
mBtnAdd_Click → CommitCurrentLine:
  1. Remove the pre-commit ShipmentsFormHydrateInvSysFromShippables call entirely
  2. Validate: NAS Inv (from mShippables(r,4)) - active qty - requested qty >= floor
  3. Write row to TABLE_SHIPMENTS
  4. Queue A/B reserve event
  5. Update mShippables(r,8) projected quantity in memory
  6. Re-render shippables listbox
```

Add should never touch `invSys.TOTAL INV`. The local `invSys.SHIPMENTS` column can be updated for floor-guard bookkeeping only via `ApplyShipmentDeltasLocal` (which correctly reads/writes only `SHIPMENTS`, not `TOTAL INV`, when `deductTotalInv = False`).

***

## Q6 — Should `ShipmentsTally` and `InventoryManagement` Be Retired from the User-Facing Form?

**Yes for user-facing visibility; retained as hidden-only backend surfaces.**

`ShipmentsTally` must continue to exist as the host sheet for `TABLE_SHIPMENTS` and `TABLE_NOTSHIPPED` — the form reads from it directly. But it should be hidden at all times from the user via `EnforceShippingSupportSheetsHidden` and **never activated by any form path**. The current call to `EnsureInventoryManagementSurface` inside `GetWritableShippingInvSysTable` is the trigger for `InventoryManagement` becoming visible — it may activate the sheet during a `ListObjects.Add` or `AutoFit` operation if the sheet was previously hidden with `xlSheetVeryHidden`. 

The fix: `EnsureInventoryManagementSurface` must never be called from display or form-open paths. `InventoryManagement!invSys` should be treated as write-only by the reservation/sent paths and read-only by the display path, accessed programmatically without sheet activation. Replace all `ws.Activate` calls in the surface-ensure code with direct `ListObject` range manipulation.

***

## Q7 — Low-Overhead Instrumentation Plan

Add a module-level timing array to `frmShipmentsTally` (or a lightweight `modShipmentsTimer` module) that writes lap times without any UI calls during loading:

```vba
' At top of frmShipmentsTally
Private mTimerLog() As String
Private mTimerCount As Long
Private mTimerStart As Single

Private Sub TLap(ByVal label As String)
    mTimerCount = mTimerCount + 1
    ReDim Preserve mTimerLog(1 To mTimerCount)
    mTimerLog(mTimerCount) = Format$(CLng((Timer - mTimerStart) * 1000), "00000") & " ms  " & label
End Sub
```

Call `TLap` at each sub-step in `InitializeFromShipping`, each control handler, and each major function entry point (`LoadShippables`, `LoadShipmentState`, `RefreshProjectedShippableInventory`, `GetWritableShippingInvSysTable`, `BoxMakerFormLoadShippableVersionInventory`, `BoxMakerShippingBomSourceTable`, any `OpenWorkbookHiddenShipping` call). Emit the full log to `ShowStatus` only on completion — zero per-step UI refresh overhead.

For `TOTAL INV` write detection, add a single guard wrapper:

```vba
Private Sub WriteToTotalInv(ByVal invLo As ListObject, ByVal rowIdx As Long, ByVal newValue As Variant, ByVal callerName As String)
    TLap "WRITE TOTAL INV row=" & rowIdx & " val=" & newValue & " from=" & callerName
    invLo.DataBodyRange.Cells(rowIdx, ColumnIndex(invLo, "TOTAL INV")).Value = newValue
End Sub
```

Replace every direct `.Cells(r, colTotal).Value =` assignment with this wrapper during the debugging phase. The timer log will immediately show which path is writing `TOTAL INV` and from where.

***

## Recommended Next Steps (Priority Order)

1. **Remove the `ShipmentsFormHydrateInvSysFromShippables` call from `CommitCurrentLine`** — this is the single highest-impact fix and directly explains 13→0. One line in `frmShipmentsTally` to delete.

2. **Remove `ReconcileShippableTotalsFromVersionInventory` from all `GetWritableShippingInvSysTable` call chains** — this explains 19→24. One call to comment out in `GetWritableShippingInvSysTable` and verify no BoxMaker regressions.

3. **Add the `TLap` instrumentation** before any further refactoring — so the next live test shows exactly which sub-step in `LoadShippables` owns the remaining 60 seconds.

4. **Block `BoxMakerShippingBomSourceTable`'s NAS workbook open from `ShipmentsFormLoadShippables`** — require `ShippingBOMView` to already be populated; surface a status message if it isn't rather than auto-rebuilding inline.

5. Split `GetWritableShippingInvSysTable` as described in Q3 and route form display through the read-only variant.

6. Make `NAS Inv` column in `mShippables` sourced exclusively from `invSysData_Shipping` or the snapshot cache — never from `invSys.TOTAL INV`, version logs, or the projected value.