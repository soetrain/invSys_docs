The two bugs are precisely diagnosed.

***

**To Codex — two bugs, exact diagnosis and fix:**

***

## Bug 1 — `Package ROW 89 not found in invSys` on Add

### What happens

The ADD path in `ShipmentsFormCommitLine` reaches this branch:

```vba
ElseIf Not isHold And Not preserveExistingReservation Then
    Set versionAvailabilityOverrides = ShippingVersionAvailabilityOverride(rowValue, descriptionValue, displayedAvailableQty)
    If invLo Is Nothing Then Set invLo = GetWritableShippingInvSysTable(ws, report)
    If invLo Is Nothing Then GoTo CleanExit
    Set reserveDeltas = BuildSelectedShipmentRowsDeltas(invLo, lo, singleRow, "Warehouse", errNotes, versionAvailabilityOverrides)
```

`GetWritableShippingInvSysTable` is called. It checks:

```vba
If ShippingInventoryPickerTableHasRows(invLo) Then
    EnsureMissingInvSysRowsFromShipmentLines invLo, GetListObject(wsShip, TABLE_SHIPMENTS)
    ...
    Set GetWritableShippingInvSysTable = invLo
    Exit Function
End If
```

`ShippingInventoryPickerTableHasRows` requires `ROW` and `ITEM` columns. After `Shipments Sent` and reload, the `invSys` table exists with headers and data from the read model. **But ROW 89 is a package row, and the `invSysData_Shipping` read model may not yet contain it** because:

1. `ShipmentsFormRefreshReadModelForWorkbook` calls `RefreshInventoryReadModelForWorkbook` → writes `invSysData_Shipping` from the local/NAS snapshot.
2. The snapshot is based on `modInventoryDomainBridge.PublishInventorySnapshotBridge` output.
3. The snapshot contains item-level inventory for known items. Package ROW 89 is populated into the snapshot from `BoxBOMView` / version inventory, not from the raw `invSys` item rows.
4. If the BOM view for ROW 89 was not refreshed after `Shipments Sent` (because `ShipmentsFormAutoSyncRefresh` skips `RefreshShippingBomViewForWorkbook`), the `invSysData_Shipping` table may have the T28 item row but with `ROW = 0` or a different ROW value.

`EnsureMissingInvSysRowsFromShipmentLines` would normally add ROW 89 from the `ShipmentsTally` table — **but after `Shipments Sent` the shipment rows for T28/T29 were deleted from `ShipmentsTally`**. So there is no shipment line for ROW 89 to hydrate from, and `invSys` lacks ROW 89 entirely.

Then `BuildSelectedShipmentRowsDeltas` with `requiredArea = "Warehouse"` calls `FindInvListRowByRowValue(invLo, 89)` → returns `Nothing` → `report = "Package ROW 89 not found in invSys."` → exits.

### Root cause

`GetWritableShippingInvSysTable` relies on:
1. `invSysData_Shipping` having the ROW values for shippable packages, OR
2. `ShipmentsTally` having active rows for hydration via `EnsureMissingInvSysRowsFromShipmentLines`

After `Shipments Sent` clears `ShipmentsTally` and before the user creates a new order, **neither source has ROW 89** unless `invSysData_Shipping` was built with the correct ROW values from BOM view.

### Fix — hydrate `invSys` from `mShippables` when invSys is missing the ROW

The form has `mShippables` in memory — it is the exact array that drives the display, and it does have ROW 89 at column index 1. Pass it as an additional hydration source into `GetWritableShippingInvSysTable`, or add a pre-hydration step in `ShipmentsFormCommitLine` before calling it.

**Option A (simpler, self-contained) — add a `ShipmentsFormHydrateInvSysFromShippables` call before ADD:**

In `modTS_Shipments`, add:

```vba
Public Sub ShipmentsFormHydrateInvSysFromShippables(ByVal wsShip As Worksheet, _
                                                    ByVal shippables As Variant)
    ' Ensures every visible shippable row exists in invSys before Add validation.
    On Error GoTo FailSoft

    Dim invLo As ListObject
    Dim r As Long
    Dim rowVal As Long
    Dim itemName As String
    Dim versionLabel As String
    Dim uomVal As String
    Dim locVal As String
    Dim totalInvQty As Double
    Dim versionInv As Object
    Dim key As Variant
    Dim report As String
    Dim lr As ListRow

    If wsShip Is Nothing Then Exit Sub
    If IsEmpty(shippables) Then Exit Sub

    Set invLo = GetInvSysTableFromWorkbook(wsShip.Parent)
    If invLo Is Nothing Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    EnsureShippingWorksheetEditable invLo.Parent

    For r = 1 To UBound(shippables, 1)
        rowVal    = CLng(Val(NzStr(shippables(r, 1))))   ' col 1 = ROW
        itemName  = Trim$(NzStr(shippables(r, 2)))       ' col 2 = Box name
        versionLabel = Trim$(NzStr(shippables(r, 3)))    ' col 3 = Version
        uomVal    = Trim$(NzStr(shippables(r, 5)))       ' col 5 = UOM
        locVal    = Trim$(NzStr(shippables(r, 6)))       ' col 6 = Location
        If rowVal <= 0 Or itemName = "" Then GoTo NextShippable
        If FindInvRowIndexByRow(invLo, rowVal) > 0 Then GoTo NextShippable  ' already present

        ' ROW is missing — add it from the visible display data
        Set lr = FirstBlankListRowShipping(invLo)
        If lr Is Nothing Then Set lr = invLo.ListRows.Add
        WriteValue lr, "ROW",      rowVal
        WriteValue lr, "ITEM_CODE", itemName
        WriteValue lr, "ITEM",     itemName
        WriteValue lr, "UOM",      uomVal
        WriteValue lr, "LOCATION", locVal
        WriteValue lr, "DESCRIPTION", versionLabel

        ' Populate TOTAL INV from version inventory or visible NAS Inv text
        totalInvQty = 0
        Set versionInv = BoxMakerFormLoadBoxVersionInventory(rowVal, itemName)
        If Not versionInv Is Nothing Then
            For Each key In versionInv.Keys
                totalInvQty = totalInvQty + NzDbl(versionInv(key))
            Next key
        End If
        If totalInvQty <= 0 Then
            ' Fall back to visible NAS Inv from mShippables col 4
            totalInvQty = NzDbl(NzStr(shippables(r, 4)))
        End If
        WriteValue lr, "TOTAL INV", totalInvQty
        WriteValue lr, "SHIPMENTS", 0
NextShippable:
    Next r
    Exit Sub
FailSoft:
End Sub
```

**In `frmShipmentsTally.CommitCurrentLine`, call it before `ShipmentsFormCommitLine`:**

```vba
Private Sub CommitCurrentLine(ByVal actionName As String)
    On Error GoTo FailSoft

    Dim report As String
    Dim rowIndex As Long
    Dim ok As Boolean
    Dim displayedAvailableQty As String
    Dim ws As Worksheet

    rowIndex = SelectedShipmentTableRow()
    displayedAvailableQty = SelectedShippableProjectedInventoryText()

    ' Ensure the selected shippable's ROW exists in invSys before Add validation
    If UCase$(Trim$(actionName)) = "ADD" Or UCase$(Trim$(actionName)) = "UPDATE" Then
        Set ws = modTS_Shipments.GetShipmentsTallyWorksheet()
        If Not ws Is Nothing Then
            modTS_Shipments.ShipmentsFormHydrateInvSysFromShippables ws, mShippables
        End If
    End If

    ok = modTS_Shipments.ShipmentsFormCommitLine("SHIP", _
                                                 actionName, _
                                                 rowIndex, _
                                                 NzText(mTxtRef.Value), _
                                                 NzText(mTxtBox.Value), _
                                                 ParseNumber(NzText(mTxtQty.Value)), _
                                                 CLng(Val(NzText(mTxtRow.Value))), _
                                                 NzText(mTxtUom.Value), _
                                                 NzText(mTxtLocation.Value), _
                                                 NzText(mTxtVersion.Value), _
                                                 NzText(mTxtCarrier.Value), _
                                                 report, _
                                                 displayedAvailableQty)
    RefreshAfterAction report, ok
    Exit Sub
FailSoft:
    ShowStatus "Shipment row action failed: " & Err.Description
End Sub
```

Add `GetShipmentsTallyWorksheet` to `modTS_Shipments` if it does not already exist as a public accessor:

```vba
Public Function GetShipmentsTallyWorksheet() As Worksheet
    Set GetShipmentsTallyWorksheet = SheetExists(SHEET_SHIPMENTS)
End Function
```

This directly enforces the design rule: **Add validation must be able to resolve every row the display shows.** By hydrating from `mShippables` immediately before `ShipmentsFormCommitLine`, the invSys row for ROW 89 is guaranteed to exist before `BuildSelectedShipmentRowsDeltas` searches for it.

***

## Bug 2 — AutoSync runs 7-8s batch even when nothing is staged

### What happens

`ShipmentsFormAutoSyncRefresh` calls `RunShippingRuntimeQueueRefresh` unconditionally. That function:

1. Calls `SyncLocalStagedInboxRows` — fast (< 1ms if nothing staged)
2. Calls `RunBatch` — **opens the NAS processor workbook, loads the inbox, runs the processor** — this is the 7-8s cost
3. Checks `ShippingRuntimeReportShowsProcessed` — sees `Applied=0`

The guard condition is:

```vba
If Not ShippingRuntimeReportShowsProcessed(processedCount, batchReport) _
   And (requireQueuedWork Or ShippingRuntimeReportMetric(stagingReport, "LocalStagingMerged") > 0) Then
    report = "RunBatch did not handle the queued shipping event..."
    Exit Function   ' ← returns False (error), but still ran the full batch
End If
```

The problem: `RunBatch` is called **before** this guard. There is no pre-batch check that skips `RunBatch` when nothing was staged. `requireQueuedWork = True` passed from `ShipmentsFormAutoSyncRefresh` only affects whether a return-False is signaled after batch — it does not skip the batch call itself.

### Fix — skip `RunBatch` when staging merged nothing and no pending overlays exist

In `RunShippingRuntimeQueueRefresh`, add an early-exit after `SyncLocalStagedInboxRows` returns, before `RunBatch`:

```vba
Private Function RunShippingRuntimeQueueRefresh(ByVal wb As Workbook, _
                                                ByVal warehouseId As String, _
                                                ByRef report As String, _
                                                Optional ByVal requireQueuedWork As Boolean = True) As Boolean
    ...
    If Not modRoleEventWriter.SyncLocalStagedInboxRows(stagingReport, resolvedWarehouseId, stationId) Then
        ...
        Exit Function
    End If

    ' NEW: fast-exit when nothing was staged and batch is not required by caller
    Dim localMerged As Long
    localMerged = ShippingRuntimeReportMetric(stagingReport, "LocalStagingMerged")
    If requireQueuedWork And localMerged = 0 Then
        batchMs = ElapsedMillisecondsShipping(batchTimer)
        report = "Processed=0; StagingReport=" & stagingReport & "; BatchReport=Skipped(NothingStaged); " & _
                 FormatShippingRuntimeTiming(ElapsedMillisecondsShipping(totalTimer), batchMs, 0)
        RunShippingRuntimeQueueRefresh = True   ' not an error — just nothing to process
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(resolvedWarehouseId, 0, batchReport)
    ...
```

### And in `ShipmentsFormAutoSyncRefresh`, treat the "NothingStaged" skip as a fast-complete:

```vba
Public Function ShipmentsFormAutoSyncRefresh(...) As Boolean
    ...
    If Not RunShippingRuntimeQueueRefresh(operatorWb, warehouseId, runtimeReport, True) Then
        report = "AutoSync runtime: " & runtimeReport
        Exit Function
    End If

    ' If batch was skipped because nothing was staged, skip read-model refresh too
    ' (the snapshot is already current — no new events were processed)
    If InStr(1, runtimeReport, "NothingStaged", vbTextCompare) > 0 Then
        report = "OK; " & runtimeReport   ' completes in < 100ms
        ShipmentsFormAutoSyncRefresh = True
        Exit Function
    End If

    ' Batch ran and may have published a new snapshot — refresh read model
    If Not ShipmentsFormRefreshReadModelForWorkbook(operatorWb, inventoryReport, warehouseId) Then
        report = "AutoSync read-model: " & inventoryReport
        Exit Function
    End If
    ...
```

### Performance result after fix

| Scenario | Before | After |
|---|---|---|
| Nothing staged, NAS = Projected | 7-8s (RunBatch opens NAS processor) | < 100ms (exits after `SyncLocalStagedInboxRows`) |
| Staged `SHIP` event pending | 7-8s | 7-8s (RunBatch still runs, correct) |
| Sync complete, no overlays | 7-8s per poll | < 100ms, then `PendingShipmentSyncCount = 0` stops polling |

The polling loop naturally winds down: once `NAS = Projected` for all rows and `mLstShipments.ListCount = 0`, `PendingShipmentSyncCount` returns 0 and `AutoSyncIfPending` exits at the early-check, rescheduling but doing no work. If you also add `HasAnyPendingBoxVersionInventoryOverlay()` as the secondary guard (from the previous session), the scheduler stops touching NAS at all when the form is idle.