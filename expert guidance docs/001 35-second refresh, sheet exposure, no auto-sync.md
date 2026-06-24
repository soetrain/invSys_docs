Now the full picture is clear. Here is the precise brief:

***

**To Codex — 35-second refresh, sheet exposure, no auto-sync: three separate issues, targeted fixes:**

***

## Issue 1 — 35-second Refresh: `OpenShippingBomWorkbook` is the bottleneck

The call chain on Refresh button:

```
mBtnRefresh_Click
  → ShipmentsFormRefreshRuntimeInventory
      → RunShippingRuntimeQueueRefresh   (RunBatch + SyncLocalStagedInboxRows)
      → RefreshShippingBomViewForWorkbook
          → OpenShippingBomWorkbook(...)   ← opens external file workbook from NAS/network share
          → EnsureShippingBomSchema(wbBom) ← may rebuild BOM schema in that workbook
          → CopyShippingBomTable(loBom, loView)  ← bulk copies rows
          → CloseWorkbookNoSaveShipping    ← closes again
  → InitializeFromShipping                ← full form reload
      → LoadShippables                    ← reads ShippingBOMView back out
```

`OpenShippingBomWorkbook` opens a file from `target.RuntimeRoot` on a network path. On a slow NAS share, that open alone can take 10–30 seconds. `RefreshShippingBomViewForWorkbook` is called **every** Refresh — even when the BOM hasn't changed and `ShippingBOMView` already has valid data.

**Fix: gate `RefreshShippingBomViewForWorkbook` behind a staleness check**

The table already has a freshness guard at the top:

```vba
Set loView = GetShippingBomViewTable(operatorWb)
If loView Is Nothing Then
    ' only enters here if table is missing — rare
```

Extend the guard to skip the network open when the view already has rows AND the caller is not a forced full-rebuild:

```vba
Public Function RefreshShippingBomViewForWorkbook(ByVal operatorWb As Workbook, _
                                                  ByRef report As String, _
                                                  Optional ByVal forceRebuild As Boolean = False) As Boolean
    Set loView = GetShippingBomViewTable(operatorWb)

    ' Fast path: BOM view exists and caller is not forcing rebuild
    If Not loView Is Nothing And Not forceRebuild Then
        If Not loView.DataBodyRange Is Nothing Then
            If loView.DataBodyRange.Rows.Count > 0 Then
                report = "BOM view already populated; skipped network open."
                RefreshShippingBomViewForWorkbook = True
                Exit Function
            End If
        End If
    End If
    ' ... existing slow path (open wbBom, copy, close) ...
```

In `ShipmentsFormRefreshRuntimeInventory`, change the call:

```vba
' Normal Refresh: skip BOM network open if view is already populated
If Not RefreshShippingBomViewForWorkbook(wb, bomReport, False) Then ...

' You can expose a "Full Rebuild" option separately for diagnostics
```

This alone will drop normal Refresh from 35 seconds to under 1 second on most runs.

Also profile `RunBatch` — if it's scanning large event logs, `RequireQueuedWork=False` is already passed from `ShipmentsFormRefreshRuntimeInventory`, meaning it runs unconditionally. If `RunBatch` itself takes 20+ seconds, add a `processedCount = 0` early-exit path when `SyncLocalStagedInboxRows` reports nothing to merge:

```vba
' In RunShippingRuntimeQueueRefresh:
If Not SyncLocalStagedInboxRows(stagingReport, ...) Then ... exit
If ShippingRuntimeReportMetric(stagingReport, "LocalStagingMerged") = 0 Then
    ' Nothing was staged, skip RunBatch — no queued work
    processedCount = 0
    batchReport = "Skipped=NothingStaged"
    RunShippingRuntimeQueueRefresh = True
    Exit Function
End If
processedCount = modProcessor.RunBatch(...)
```

***

## Issue 2 — Sheet visibility: `BeginShippingTableMutation` makes sheets visible and doesn't always restore them

`BeginShippingTableMutation` sets `lo.Parent.Visible = xlSheetVisible` (line 8574) to allow table writes, and `EndShippingTableMutation` restores `previousVisibility`. This is correct in isolation — but if any code path exits early (error, `GoTo CleanExit`) **without** calling `EndShippingTableMutation`, the sheet stays visible.

Additionally, `EnsureShippingWorkbookSurface` (called in `InitializeShippingModule` at line 127) may leave surfaces visible if it creates them. There is no unconditional sheet-visibility enforcement after the full refresh completes.

**Fix: enforce visibility at the end of `InitializeFromShipping`**

Add an unconditional sheet-hiding pass as the last step of every form load and Refresh:

```vba
' In modTS_Shipments — new public sub:
Public Sub EnforceShippingSupportSheetsHidden(ByVal wb As Workbook)
    Const SUPPORT_SHEETS As String = "ShipmentsTally|InventoryManagement|ShippingBOM"
    Dim names() As String
    Dim i As Long
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    names = Split(SUPPORT_SHEETS, "|")
    For i = 0 To UBound(names)
        On Error Resume Next
        Set ws = wb.Worksheets(Trim$(names(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ' Cannot hide if it's the only visible sheet
            If wb.Worksheets.Count > 1 Then ws.Visible = xlSheetVeryHidden
        End If
    Next i
End Sub
```

Call it at the end of `InitializeFromShipping`, after `mLoading = False`:

```vba
' In frmShipmentsTally.InitializeFromShipping, in CleanExit block:
CleanExit:
    On Error Resume Next
    modTS_Shipments.EnforceShippingSupportSheetsHidden ActiveWorkbook   ' ← add this
    If quietStarted Then modUiQuiet.EndQuietUi
    Me.MousePointer = previousPointer
    On Error GoTo 0
    Exit Sub
```

Also call it at the end of `RefreshAfterAction` in the form:

```vba
Private Sub RefreshAfterAction(ByVal report As String, ByVal ok As Boolean)
    ...
    LoadShipmentState
    RefreshProjectedShippableInventory
    modTS_Shipments.EnforceShippingSupportSheetsHidden ActiveWorkbook   ' ← add
End Sub
```

And after `EndShippingTableMutation` in any mutation that handles `TABLE_NOTSHIPPED` or other support tables that become temporarily visible.

***

## Issue 3 — No auto-sync: the form has no polling timer

There is no `mTimer`, `OnTime`, or `Application.OnTime` reference anywhere in the form. The sync state label (`UpdateSyncStateLabel`) is only updated when a user action fires. The form has no background update mechanism.

**Fix: lightweight `Application.OnTime` poll**

The poll must be **cheap** — no `RunBatch`, no file opens. It should only check whether the snapshot/read-model NAS values have changed since the last load.

```vba
' In frmShipmentsTally:
Private mNextPollTime As Date
Private Const POLL_INTERVAL_SECONDS As Long = 15

Public Sub ScheduleAutoSync()
    On Error Resume Next
    mNextPollTime = Now + TimeSerial(0, 0, POLL_INTERVAL_SECONDS)
    Application.OnTime mNextPollTime, "modTS_Shipments.TriggerShipmentsFormAutoSync", , True
End Sub

Public Sub CancelAutoSync()
    On Error Resume Next
    If mNextPollTime > 0 Then
        Application.OnTime mNextPollTime, "modTS_Shipments.TriggerShipmentsFormAutoSync", , False
    End If
End Sub
```

In `modTS_Shipments`, add the callback (must be a Public Sub in a standard module for `OnTime`):

```vba
Public Sub TriggerShipmentsFormAutoSync()
    ' Cheap poll: check snapshot timestamp / pending count only
    Dim form As frmShipmentsTally
    ' Get the open form instance — use whatever the project's pattern is
    If Not gShipmentsTallyForm Is Nothing Then
        If gShipmentsTallyForm.Visible Then gShipmentsTallyForm.AutoSyncIfPending
    End If
End Sub
```

In `frmShipmentsTally`, add `AutoSyncIfPending`:

```vba
Public Sub AutoSyncIfPending()
    ' Only do work if sync is actually pending
    If Not modTS_Shipments.ShipmentsHavePendingLocalEvents() Then
        UpdateSyncStateLabel
        ScheduleAutoSync   ' reschedule regardless
        Exit Sub
    End If

    ' Cheap read-model check: did NAS snapshot change since last load?
    Dim newInv As Variant
    newInv = modTS_Shipments.ShipmentsFormLoadShippables()   ' reads local cached snapshot — fast
    If modTS_Shipments.ShippablesInventoryChanged(mShippables, newInv) Then
        mShippables = newInv
        PreserveMissingShippableInventory CurrentShippableInventoryCache()
        modTS_Shipments.EvictCompletedShipmentInventoryOverlaysForShippables mShippables
        RefreshProjectedShippableInventory
    End If
    UpdateSyncStateLabel
    ScheduleAutoSync
End Sub
```

`ShipmentsHavePendingLocalEvents` is a thin wrapper that checks the pending overlay/event count without touching the file system:

```vba
Public Function ShipmentsHavePendingLocalEvents() As Boolean
    ShipmentsHavePendingLocalEvents = (PendingShippingEventCount() > 0 Or HasAnyPendingBoxVersionInventoryOverlay())
End Function
```

Wire into the form lifecycle:

```vba
' In UserForm_Activate or at the end of InitializeFromShipping:
ScheduleAutoSync

' In UserForm_Deactivate / Hide / Close:
CancelAutoSync
```

**Important**: `ShipmentsFormLoadShippables` calls `BoxMakerFormLoadShippableVersionInventory` which calls `RefreshShippingBomViewForWorkbook`. After Fix 1 (BOM fast-path gate), this becomes cheap because the BOM view is already populated. If `BoxMakerFormLoadShippableVersionInventory` has its own NAS open call for the inventory snapshot specifically, add the same populated-table guard there.

***

## Instrumentation (add before optimizing further)

Add step timings to `InitializeFromShipping`:

```vba
Public Sub InitializeFromShipping()
    ...
    Dim t0 As Single: t0 = Timer
    LoadCarrierChoices
    DebugShipping "InitializeFromShipping: LoadCarrierChoices " & ElapsedMilliseconds(t0) & "ms"

    Dim t1 As Single: t1 = Timer
    LoadShippables
    DebugShipping "InitializeFromShipping: LoadShippables " & ElapsedMilliseconds(t1) & "ms"

    Dim t2 As Single: t2 = Timer
    LoadShipmentState
    DebugShipping "InitializeFromShipping: LoadShipmentState " & ElapsedMilliseconds(t2) & "ms"

    Dim t3 As Single: t3 = Timer
    RefreshProjectedShippableInventory
    DebugShipping "InitializeFromShipping: RefreshProjectedShippableInventory " & ElapsedMilliseconds(t3) & "ms"
    ...
```

And in `ShipmentsFormRefreshRuntimeInventory`:

```vba
Dim t0 As Single: t0 = Timer
RunShippingRuntimeQueueRefresh ...
DebugShipping "RefreshRuntime: RunBatch " & ElapsedMillisecondsShipping(t0) & "ms"

Dim t1 As Single: t1 = Timer
RefreshShippingBomViewForWorkbook ...
DebugShipping "RefreshRuntime: BOMRefresh " & ElapsedMillisecondsShipping(t1) & "ms"
```

`DebugShipping` writes to `Debug.Print` or a local log — use whatever pattern is already established. This will immediately show whether the 35 seconds is in `RunBatch`, `OpenShippingBomWorkbook`, or somewhere else.

***

## Priority order

1. **BOM fast-path gate** — eliminates the 35-second network open for normal Refresh. One `If Not forceRebuild` guard.
2. **`EnforceShippingSupportSheetsHidden`** — called from `CleanExit` in `InitializeFromShipping` and `RefreshAfterAction`. Fixes sheet exposure permanently.
3. **`Application.OnTime` auto-poll** — 15-second interval, cheap snapshot check. Fixes the "waited a minute, nothing happened" UX.
4. **Step timings** — confirm which step is slow before doing further structural work.