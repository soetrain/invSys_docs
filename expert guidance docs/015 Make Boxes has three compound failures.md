Now I have a complete picture of the bug. I'm ready to write the full diagnosis and fix guidance.

***

Here is a precise, actionable breakdown of the defect and the changes needed to fix it.

## Root Cause Analysis

The symptom after `Make Boxes` for **T31 v1 (row 92, qty 10)** is:

- `NAS Inv: unknown`, `Projected Inv: 10`, sync gauge stays red.

There are **three compounding failures**, all in the finished-box version inventory path.

***

## Failure 1 — `BuildBoxVersionInventoryCache` never populates version totals because `tblInventoryLog` is empty in the operator workbook

`BoxMakerFormLoadShippableVersionInventory` (line ~4765) calls `BuildBoxVersionInventoryCache`, which reads `tblInventoryLog` from the **operator workbook** (not the NAS inventory workbook) and scans for `BOX_BUILD` rows containing `VERSION=v1` in the `Note` column.

`RefreshOperatorInventoryLogForWorkbook` is supposed to copy the server log into the operator workbook each autosync cycle. The popup evidence says:

```
SnapshotError=Snapshot workbook not resolved.
PublishWarning=Snapshot workbook not resolved.
Refresh=0
```

This means the snapshot workbook (which `modInventoryDomainBridge.ResolveInventoryWorkbookBridge` uses) is not resolving. `RefreshOperatorInventoryLogForWorkbook` silently returns with `"InventoryLog refresh skipped: inventory workbook not resolved."` — and because `ShipmentsFormAutoSyncRefresh` suppresses non-"OK" `logReport` strings, the failure is invisible in the status bar. The operator `tblInventoryLog` never receives the BOX_BUILD row, so `BuildBoxVersionInventoryCache` finds nothing and returns an empty version-totals dictionary.

**Fix:** Add explicit log-report surfacing in `ShipmentsFormAutoSyncRefresh` regardless of whether it equals "OK", and resolve why the snapshot workbook path fails. Specifically, around line 6510:

```vba
' CURRENT (suppresses non-OK logReport silently):
If Trim$(logReport) <> "" And StrComp(Trim$(logReport), "OK", vbTextCompare) <> 0 Then
    report = report & "; " & logReport
End If

' CHANGE TO (always surface logReport for Box Maker debugging):
If Trim$(logReport) <> "" Then
    report = report & "; LOG:" & logReport
End If
```

Also verify `CurrentShippingInventoryWorkbookPath(warehouseId)` returns the correct NAS path on your Zenbook and that `modInventoryDomainBridge.ResolveInventoryWorkbookBridge` can open it. The "Snapshot workbook not resolved" error is a separate but related path — it is likely that the inventory workbook itself (`*.invSys.Data.Inventory.xlsb`) is the file that fails, not only the snapshot, because `RefreshOperatorInventoryLogForWorkbook` calls that same bridge.

***

## Failure 2 — Multi-version box skips `ResolveCurrentInventoryValue` for `rowData(4)`

In `BoxMakerFormLoadShippableVersionInventory` (lines ~4783–4798):

```vba
' Only resolves row-level current inventory when there is exactly ONE active version:
If activeCount = 1 Then
    foundCurrent = False
    currentInv = ResolveCurrentInventoryValue(ws, invLo, boxRow, ...)
    If foundCurrent Then rowData(4) = currentInv
End If
' Then falls through to version-specific cache — which is empty (Failure 1):
If Not versionInv Is Nothing Then
    If versionInv.Exists(versionLabel) And Trim$(NzStr(rowData(4))) = "" Then rowData(4) = versionInv(versionLabel)
End If
```

Because T31 has **two active versions** (`activeCount = 2`), `rowData(4)` is **never filled from `invSys` TOTAL INV** — by design to avoid mixing v1 and v2 totals. That is architecturally correct. But it means the only way `rowData(4)` can become non-blank is through `versionInv`, which is empty (Failure 1). Until the log copy is fixed, `rowData(4)` will always be empty.

**Note:** This is not a code bug — it is the correct multi-version guard. Fix Failure 1 first.

***

## Failure 3 — `mPendingVersionInv` in the form is never evicted because the backend `rowData(4)` stays blank

`DisplayBoxVersionInventoryText` (form, line ~1252) only removes the key from `mPendingVersionInv` when:

```vba
If Abs(backendQty - pendingQty) < 0.0000001 Then
    mPendingVersionInv.Remove key
```

Since `rowData(4)` is blank, `backendQty = 0`, never matching `pendingQty = 10`. The key persists. `PendingShippableInventoryCount()` counts it, keeping the gauge red indefinitely.

Same problem exists in the persistent overlay in `modTS_Shipments`. `PendingBoxVersionInventoryOverlayValue` (line ~5262) only evicts when `backendQty > baselineQty + 0.0000001`. `baselineQty` was saved as 0 (since NAS was unknown at post time), and `backendQty` from the cache is also 0, so the condition `backendQty > 0 + 0.0000001` is never true.

There is a secondary clearing path using the `"SENT|"` prefix key — `EvictCompletedShipmentInventoryOverlaysForShippables` — but it also requires `backendQty >= overlayQty` (qty 10), which never happens because the backend is blank.

**Fix — two changes:**

### Fix 3a — `modTS_Shipments.EvictCompletedShipmentInventoryOverlaysForShippables`

The eviction condition is inverted — it should remove the overlay when the **backend is numeric AND equals or exceeds the overlay**, but it currently also never fires when backend is blank/unknown. Add a "staleness" fallback: if the event `Applied=1` was confirmed in the last sync (traceable from the `batchProcessed` boolean already in `CommitBoxMakerFormAction`), promote the overlay from `"active"` key to `"SENT|"` key. This is already architecturally set up via `RegisterPendingBoxVersionInventoryOverlay` → `PendingBoxVersionInventoryKey` vs `SentPendingBoxVersionInventoryKey`, but nothing promotes the key from pending to sent.

Look at line ~5360 in `modTS_Shipments.bas`:

```vba
Public Sub PromoteActiveOverlayToSent(ByVal packageRow As Long, ByVal versionLabel As String)
```

Check if this sub exists (the grep found it at line 5360). If it does, call it from `frmShippingBoxMaker.AutoSyncIfPending` **when `ShipmentsFormAutoSyncRefresh` returns True and `syncCompleted` was True in `CommitBoxMakerFormAction`**:

In `frmShippingBoxMaker`:

```vba
' After ShipmentsFormAutoSyncRefresh returns True, promote overlay to SENT:
If modTS_Shipments.HasSentOverlayForRowVersion(mSelectedPackageRow, SelectedVersionLabel()) = False Then
    modTS_Shipments.PromoteActiveOverlayToSent mSelectedPackageRow, SelectedVersionLabel()
End If
```

### Fix 3b — Form-level `mPendingVersionInv` bypass when version NAS resolves via `versionInv`

When `AutoSyncIfPending` calls `RefreshShippableInventoryCache True` and then `RenderShippableInventoryFromCache`, the render path calls `DisplayBoxVersionInventoryText` with the new `rowData(4)` from cache. Once Failure 1 is fixed and the log copy works, `rowData(4)` will be numeric and the existing `mPendingVersionInv.Remove key` line will fire normally.

So Fix 3b is contingent on Fix 1. It does not require separate code if Fix 1 is solved.

***

## Summary of Changes

| Priority | File | Change |
|---|---|---|
| 1 (root) | `modTS_Shipments.bas` `RefreshOperatorInventoryLogForWorkbook` | Debug why `ResolveInventoryWorkbookBridge` returns Nothing on Zenbook; verify `CurrentShippingInventoryWorkbookPath` resolves the correct NAS share path |
| 1 (root) | `modTS_Shipments.bas` `ShipmentsFormAutoSyncRefresh` | Surface `logReport` unconditionally to make the log refresh failure visible in the status bar |
| 2 | `modTS_Shipments.bas` `BuildBoxVersionInventoryCache` | Add a `Debug.Print` or temporary status log showing: candidates tried, `tblInventoryLog` row count found, matched rows, extracted version labels — to confirm the log copy landed and is being parsed |
| 3 | `frmShippingBoxMaker` `AutoSyncIfPending` | After `ShipmentsFormAutoSyncRefresh` returns `True`, call `modTS_Shipments.PromoteActiveOverlayToSent(mSelectedPackageRow, SelectedVersionLabel())` so the overlay transitions from active-pending to SENT state and becomes eligible for eviction even when backend quantity is still blank |
| 4 | `modTS_Shipments.bas` `EvictCompletedShipmentInventoryOverlaysForShippables` | Review eviction condition — currently `backendQty <= overlayQty` evicts (removes), but the guard fires only when the sent-key exists, which requires step 3 above |

The **key caution from the report remains valid**: do not use row-level `TOTAL INV` as the version quantity when multiple active versions exist. The version-specific log evidence path is architecturally correct. The only thing broken is that the log evidence never arrives in the operator workbook because the inventory workbook bridge fails silently.

***

## Quickest Diagnostic Step

Before touching any code, run this after a `Make Boxes` action and capture the output in the VBE immediate window:

```vba
' Paste into Immediate Window after clicking Make Boxes:
? modTS_Shipments.BoxMakerFormCurrentInventoryDebugReport(92, "v1")
```

If that sub doesn't exist yet, add it temporarily:

```vba
Public Function BoxMakerFormCurrentInventoryDebugReport(ByVal packageRow As Long, ByVal versionLabel As String) As String
    Dim warehouseId As String
    Dim inventoryPath As String
    Dim wb As Workbook
    Dim lo As ListObject
    Dim report As String

    warehouseId = ResolveCurrentShippingWarehouseId()
    inventoryPath = CurrentShippingInventoryWorkbookPath(warehouseId)
    report = "WarehouseId=" & warehouseId & "; InventoryPath=" & inventoryPath
    Set wb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(warehouseId)
    If wb Is Nothing Then
        report = report & "; InventoryWb=NOT RESOLVED"
    Else
        Set lo = FindListObjectByNameShipping(wb, "tblInventoryLog")
        If lo Is Nothing Then
            report = report & "; tblInventoryLog=NOT FOUND"
        Else
            report = report & "; tblInventoryLog rows=" & IIf(lo.DataBodyRange Is Nothing, "0", CStr(lo.DataBodyRange.Rows.Count))
        End If
    End If
    BoxMakerFormCurrentInventoryDebugReport = report
End Function
```

That output will immediately confirm whether the inventory workbook bridge is the culprit, without any UI changes. If `InventoryWb=NOT RESOLVED` appears, the fix is entirely in how `modInventoryDomainBridge` resolves the path — likely a NAS share path difference between your desktop (where the snapshot path is set) and the Zenbook (where you're running the live test).