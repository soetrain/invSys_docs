The root cause is now definitively confirmed. Here is the precise brief:

***

**To Codex — NAS never catches up after Shipments Sent: confirmed root cause, one-function fix:**

***

## Root cause (exact)

`AutoSyncIfPending` calls `ShipmentsFormRefreshReadModelForWorkbook`. That function calls only `modOperatorReadModel.RefreshInventoryReadModelForWorkbook` — it reads the **already-published snapshot** from local/NAS storage and writes it into the operator workbook's `invSys` table.

It does **not** call `RunShippingRuntimeQueueRefresh`. That function is only called by `ShipmentsFormRefreshRuntimeInventoryCore`, which is the manual Refresh button path.

`RunShippingRuntimeQueueRefresh` is what:
1. Calls `SyncLocalStagedInboxRows` — flushes the locally staged `SHIP` event to the NAS inbox
2. Calls `RunBatch` — processes the inbox event and applies the `SHIP` deduction to NAS inventory
3. Calls `PublishInventorySnapshotBridge` — writes the new authoritative snapshot

Without step 1–3, `RefreshInventoryReadModelForWorkbook` reads the pre-shipment snapshot forever, always returning `NAS Inv = 19`. The auto-sync loop correctly reads and re-reads that same stale snapshot reporting "OK" every 15 seconds.

Manual Refresh works because `mBtnRefresh_Click` calls `ShipmentsFormRefreshRuntimeInventory` → `ShipmentsFormRefreshRuntimeInventoryCore` → `RunShippingRuntimeQueueRefresh` + `RefreshInventoryReadModelForWorkbook` in sequence.

***

## The fix: `AutoSyncIfPending` must call the runtime queue path, not just the read-model path

Replace `ShipmentsFormRefreshReadModelForWorkbook` with `ShipmentsFormRefreshRuntimeInventoryCore` in `AutoSyncIfPending`:

```vba
' CURRENT (stale-snapshot loop):
If modTS_Shipments.ShipmentsFormRefreshReadModelForWorkbook(operatorWb, report) Then
    mLoading = True
    changedLoading = True
    LoadShippables
    LoadShipmentState
    RefreshProjectedShippableInventory
    mLoading = False
    changedLoading = False
    UpdateSyncStateLabel
Else
    UpdateSyncStateLabel
End If

' FIXED (processes queued events, then reads updated snapshot):
If modTS_Shipments.ShipmentsFormAutoSyncRefresh(operatorWb, report) Then
    mLoading = True
    changedLoading = True
    LoadShippables
    LoadShipmentState
    RefreshProjectedShippableInventory
    mLoading = False
    changedLoading = False
    UpdateSyncStateLabel
Else
    UpdateSyncStateLabel
End If
```

Add `ShipmentsFormAutoSyncRefresh` to `modTS_Shipments` — it runs the runtime queue and then the read-model refresh, but **skips the BOM network open** (already gated in `RefreshShippingBomViewForWorkbook` with `forceRebuild:=False`):

```vba
Public Function ShipmentsFormAutoSyncRefresh(ByVal operatorWb As Workbook, _
                                             ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim warehouseId As String
    Dim runtimeReport As String
    Dim inventoryReport As String

    If operatorWb Is Nothing Then
        report = "No operator workbook for auto-sync."
        Exit Function
    End If

    warehouseId = ResolveCurrentShippingWarehouseId()

    ' Step 1: flush staged inbox + run processor + publish snapshot
    If Not RunShippingRuntimeQueueRefresh(operatorWb, warehouseId, runtimeReport, False) Then
        report = "AutoSync runtime: " & runtimeReport
        Exit Function
    End If

    ' Step 2: read updated snapshot into operator invSys table
    If Not ShipmentsFormRefreshReadModelForWorkbook(operatorWb, inventoryReport, warehouseId) Then
        report = "AutoSync read-model: " & inventoryReport
        Exit Function
    End If

    report = "OK"
    If Trim$(runtimeReport) <> "" Then report = report & "; " & runtimeReport
    If Trim$(inventoryReport) <> "" And StrComp(Trim$(inventoryReport), "OK", vbTextCompare) <> 0 Then
        report = report & "; " & inventoryReport
    End If
    ShipmentsFormAutoSyncRefresh = True
    Exit Function

FailSoft:
    report = "AutoSync refresh failed: " & Err.Description
End Function
```

This is exactly `ShipmentsFormRefreshRuntimeInventoryCore` minus the BOM refresh (which is already cheap after the fast-path gate, but skipping it entirely for auto-sync is fine since the BOM doesn't change during a shipping session).

***

## Why the status message confirms this

```
AutoSync: NAS was 19, now 19. OK
```

`report = "OK"` is returned by `ShipmentsFormRefreshReadModelForWorkbook` when `RefreshInventoryReadModelForWorkbook` succeeds — it reads the snapshot cleanly. `NAS was 19, now 19` means the snapshot on disk still has `19`. The snapshot is stale because `RunBatch` was never called to process the queued `SHIP` event and republish.

After the fix, the status will show:
```
AutoSync: NAS was 19, now 18. OK; Processed=1; StagingReport=...; BatchReport=Applied=1
```

***

## Performance note for auto-sync

`RunShippingRuntimeQueueRefresh` with `requireQueuedWork:=False` runs `RunBatch` unconditionally. For auto-sync, pass `requireQueuedWork:=True` (the default) so it exits early if there is nothing staged:

```vba
' In ShipmentsFormAutoSyncRefresh:
If Not RunShippingRuntimeQueueRefresh(operatorWb, warehouseId, runtimeReport, True) Then
    ' True = requireQueuedWork: exits fast if nothing to process
    ' This path returns False only on error, not on "nothing to do"
    ' Check if it was a "nothing staged" skip vs real failure:
    If InStr(1, runtimeReport, "NothingStaged", vbTextCompare) > 0 Or _
       InStr(1, runtimeReport, "Skipped", vbTextCompare) > 0 Then
        ' Nothing queued — read-model is already current, no need to refresh
        report = "AutoSync: nothing pending."
        ShipmentsFormAutoSyncRefresh = True   ' not an error
        Exit Function
    End If
    report = "AutoSync runtime: " & runtimeReport
    Exit Function
End If
```

This keeps the 15-second poll cheap when there is nothing to process — it exits after `SyncLocalStagedInboxRows` confirms zero staged rows, without opening the NAS processor workbook or running `RunBatch`.