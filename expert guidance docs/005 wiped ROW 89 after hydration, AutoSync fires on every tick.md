**To Codex — two confirmed root causes, minimal targeted fixes:**

***

## Root Cause 1 — AutoSync fires on every tick even when idle

### Why it still runs despite `Sync: complete`

The guard in `AutoSyncIfPending` at the `syncCount <= 0` branch evicts overlays and re-checks `HasAnyPendingBoxVersionInventoryOverlay()`. That function calls `EnsurePendingBoxVersionInventoryOverlayLoaded()`, which loads from the persisted overlay file on disk. After `Shipments Sent` + NAS catch-up, the SENT overlay was evicted from the in-memory dictionary — but **`PersistPendingBoxVersionInventoryOverlay` may have written the eviction to disk, or the file still has old SENT keys from a previous session** (because `EvictCompletedShipmentInventoryOverlaysForShippables` calls `PersistPendingBoxVersionInventoryOverlay` only when `changed = True`, and the NAS backstop `backendQty <= overlayQty` condition may not have been triggered yet at first reload).

So `HasAnyPendingBoxVersionInventoryOverlay()` returns `True` despite the visible form showing `Sync: complete`, because the persisted file still has a SENT key. AutoSync falls through to `ShipmentsFormAutoSyncRefresh` → `RunShippingRuntimeQueueRefresh` → `RunBatch` → 14 seconds of wasted work.

### Confirmed design violation

`ScheduleAutoSync` is called from `UserForm_Activate` unconditionally. That means on every form open/activate, a 15-second timer fires. The form has no `mAutoSyncArmed` flag to distinguish "just sent something, need catch-up" from "idle, everything already synced."

### Fix — `mAutoSyncArmed` flag with three-state logic

**In `frmShipmentsTally.frm`, add one module-level variable:**

```vba
Private mAutoSyncArmed As Boolean
```

**`UserForm_Activate` — do not schedule auto-sync on activate:**

```vba
Private Sub UserForm_Activate()
    modTS_Shipments.RegisterShipmentsFormAutoSync Me
    ' Do NOT call ScheduleAutoSync here — armed only by Shipments Sent
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me
        mResizeInitialized = True
    End If
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub
```

**`InitializeFromShipping` — schedule only if there is actually pending work to watch:**

```vba
' Replace the final ScheduleAutoSync call in InitializeFromShipping:
    UpdateSyncStateLabel
    mLoading = False
    ...
    ' Arm and schedule only if projected != NAS on load (i.e. prior sent is pending catch-up)
    mAutoSyncArmed = (PendingShipmentSyncCount() > 0) Or modTS_Shipments.HasAnyPendingBoxVersionInventoryOverlay()
    If mAutoSyncArmed Then ScheduleAutoSync
```

**`AutoSyncIfPending` — hard exit if not armed:**

```vba
Public Sub AutoSyncIfPending()
    On Error GoTo CleanExit

    Dim operatorWb As Workbook
    Dim report As String
    Dim changedLoading As Boolean
    Dim nasBeforeRefresh As String
    Dim nasAfterRefresh As String

    If Not mAutoSyncArmed Then Exit Sub   ' ← not armed: do not reschedule, just exit
    If mLoading Then
        ShowStatus "AutoSync: skipped (loading)."
        GoTo CleanExit
    End If

    ' Re-check whether there is still anything to do
    Dim syncCount As Long
    syncCount = PendingShipmentSyncCount()
    If syncCount <= 0 Then
        EvictOrphanedActiveOverlays
        modTS_Shipments.EvictCompletedShipmentInventoryOverlaysForShippables mShippables
        RefreshProjectedShippableInventory
        syncCount = PendingShipmentSyncCount()
    End If

    If syncCount <= 0 And Not modTS_Shipments.HasAnyPendingBoxVersionInventoryOverlay() Then
        ' Nothing left to sync — disarm
        mAutoSyncArmed = False
        UpdateSyncStateLabel
        Exit Sub   ' ← disarmed: do not reschedule
    End If

    nasBeforeRefresh = FirstShippableNasText()
    Set operatorWb = ResolveOperatorWorkbook()
    If operatorWb Is Nothing Then
        ShowStatus "AutoSync: operator workbook not resolved."
        GoTo CleanExit   ' reschedule below — transient failure
    End If

    If modTS_Shipments.ShipmentsFormAutoSyncRefresh(operatorWb, report) Then
        mLoading = True
        changedLoading = True
        LoadShippables
        nasAfterRefresh = FirstShippableNasText()
        LoadShipmentState
        RefreshProjectedShippableInventory
        mLoading = False
        changedLoading = False
        UpdateSyncStateLabel
        ShowStatus "AutoSync: NAS was " & IIf(nasBeforeRefresh = "", "unknown", nasBeforeRefresh) & _
                   ", now " & IIf(nasAfterRefresh = "", "unknown", nasAfterRefresh) & ". " & report
    Else
        ShowStatus "AutoSync: refresh failed. " & report
        UpdateSyncStateLabel
    End If

CleanExit:
    If changedLoading Then mLoading = False
    If mAutoSyncArmed Then ScheduleAutoSync   ' ← only reschedule if still armed
End Sub
```

**Arm AutoSync from the form's Shipments Sent result handler.**

`RunShippingAction` already calls `LoadShipmentState` and `RefreshProjectedShippableInventory` after the sent path. Add arming there:

```vba
' In RunShippingAction, inside the ShipmentsSent result block, after RefreshProjectedShippableInventory:
If stageOnly = False Then   ' this is the Shipments Sent branch, not To Shipments
    mAutoSyncArmed = True
    ScheduleAutoSync
End If
```

Locate the exact call site: `RunShippingAction` receives a boolean `stageOnly` (or similar). The Shipments Sent button calls `RunShippingAction False`. Add:

```vba
' After RefreshAfterAction report, ok (in the btn handler or RunShippingAction post-processing):
If ok And Not stageOnly Then   ' Shipments Sent completed
    mAutoSyncArmed = True
    If Not mAutoSyncArmed Then ScheduleAutoSync  ' arm and start timer
    ScheduleAutoSync
End If
```

Or simpler — expose `ArmAutoSync` as a public method and call it from `RefreshAfterAction` when the report contains "Shipments sent":

```vba
Public Sub ArmAutoSync()
    mAutoSyncArmed = True
    ScheduleAutoSync
End Sub
```

And in `RefreshAfterAction`:

```vba
Private Sub RefreshAfterAction(ByVal report As String, ByVal ok As Boolean)
    ...
    If ok And InStr(1, report, "Shipments sent", vbTextCompare) > 0 Then ArmAutoSync
End Sub
```

***

## Root Cause 2 — `GetWritableShippingInvSysTable` wipes ROW 89 after hydration

### Confirmed sequence

`CommitCurrentLine` calls `ShipmentsFormHydrateInvSysFromShippables` → inserts ROW 89 into `invLo`.

Then `ShipmentsFormCommitLine` → ADD branch → `GetWritableShippingInvSysTable`:

```vba
If ShippingInventoryPickerTableHasRows(invLo) Then
    EnsureMissingInvSysRowsFromShipmentLines invLo, loShip
    ...
    Set GetWritableShippingInvSysTable = invLo
    Exit Function   ' ← exits here if invLo has rows
End If
' Only reaches HydrateInvSysFromShippingReadModel if invLo has NO rows
```

After hydration, `invLo` has rows (`ShippingInventoryPickerTableHasRows` returns `True`). So `GetWritableShippingInvSysTable` takes the first branch, calls `EnsureMissingInvSysRowsFromShipmentLines` (which only adds rows from `ShipmentsTally` — and `ShipmentsTally` is empty post-Sent), then returns `invLo`. **The hydrated ROW 89 should still be there** if only the first branch is taken.

But the actual wipe happens earlier — inside `LoadShippables` → `ShipmentsFormLoadShippables` → `BoxMakerFormLoadShippableVersionInventory`. That path, triggered by AutoSync, calls `GetWritableShippingInvSysTable` with a clear/reload path that calls `HydrateInvSysFromShippingReadModel` → `ClearListObjectData invLo` → wipes all rows including ROW 89. This AutoSync-driven reload happens concurrently with (or just before) `CommitCurrentLine`.

More precisely: AutoSync fires at 15-second intervals. When the user clicks Add, AutoSync may have just run `LoadShippables` which calls `ShipmentsFormLoadShippables` → which internally calls `GetWritableShippingInvSysTable` → which in the post-Sent state (no `invSysData_Shipping` rows for this box yet) reaches `HydrateInvSysFromShipmentLines` with an empty `ShipmentsTally` → `ClearListObjectData invLo` → **all rows wiped**. When Add then runs, `invLo` has no rows, hydration adds ROW 89, `GetWritableShippingInvSysTable` takes the first branch (now has rows), `EnsureMissingInvSysRowsFromShipmentLines` does nothing (ShipmentsTally empty), and ROW 89 survives.

Wait — that would make Add succeed. The actual failure is: `invLo` has rows **without ROW 89** (from a previously successful `HydrateInvSysFromShippingReadModel` that wrote `invSysData_Shipping` rows for other items but not for the package box with ROW 89). `ShippingInventoryPickerTableHasRows` returns `True` (there are rows, just not ROW 89). So `GetWritableShippingInvSysTable` exits early via the first branch without calling `HydrateInvSysFromShippingReadModel` — and the pre-commit hydration `ShipmentsFormHydrateInvSysFromShippables` added ROW 89, but **`GetWritableShippingInvSysTable` returned the correct `invLo` that still has ROW 89**. So Add should work.

Unless: `ShipmentsFormHydrateInvSysFromShippables` is not yet deployed (it was proposed last session but may not be in the live code). Check:

```vba
' In modTS_Shipments:
curl grep "ShipmentsFormHydrateInvSysFromShippables"
```

If that function is absent, the entire pre-commit hydration step doesn't exist — ROW 89 is never added — and the first branch of `GetWritableShippingInvSysTable` returns `invLo` with rows but without ROW 89 → `BuildSelectedShipmentRowsDeltas` fails.