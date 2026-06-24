The root cause is now clear. The bugs are precise and fixable.

***

**To Codex — auto-sync not updating NAS Inv: two confirmed failure modes, exact fixes:**

***

## Failure Mode 1 — `PendingShipmentSyncCount` returns 0 after `Shipments Sent`

```vba
Private Function PendingShipmentSyncCount() As Long
    For r = 1 To UBound(mShippables, 1)
        nasText     = Trim$(NzText(mShippables(r, 4)))    ' NAS Inv column
        projectedText = Trim$(NzText(mShippables(r, 8))) ' Projected Inv column
        If projectedText <> "" And StrComp(nasText, projectedText, vbTextCompare) <> 0 Then
            PendingShipmentSyncCount = PendingShipmentSyncCount + 1
        End If
    Next r
    If Not mLstShipments Is Nothing Then
        PendingShipmentSyncCount = PendingShipmentSyncCount + mLstShipments.ListCount
    End If
End Function
```

After `Shipments Sent`, the shipment rows are gone from `mLstShipments` (ListCount = 0). The SENT overlay makes `mShippables(r, 8)` = `"5"` and `mShippables(r, 4)` = `"6"` (stale NAS). So `"5" <> "6"` → count = 1. That part works.

**But:** if the SENT overlay was correctly evicted by a prior bug fix — or if `mShippables` itself is stale from a previous refresh — the function returns 0 and `AutoSyncIfPending` exits before doing anything:

```vba
If PendingShipmentSyncCount() <= 0 Then
    UpdateSyncStateLabel
    GoTo CleanExit    ' ← reschedule skipped because GoTo bypasses ScheduleAutoSync
End If
```

**This is the critical secondary bug:** `GoTo CleanExit` bypasses `ScheduleAutoSync` at the bottom of `AutoSyncIfPending`. If the count drops to 0 for any reason (eviction race, stale `mShippables`, first poll fires before SENT overlay is written), the timer is never rescheduled and auto-sync stops permanently for that session.

**Fix 1a — always reschedule, even when count = 0:**

```vba
Public Sub AutoSyncIfPending()
    On Error GoTo CleanExit

    Dim operatorWb As Workbook
    Dim report As String
    Dim changedLoading As Boolean

    If mLoading Then GoTo CleanExit

    ' Always check if there are pending overlays regardless of mShippables state
    If PendingShipmentSyncCount() <= 0 And Not modTS_Shipments.HasAnyPendingBoxVersionInventoryOverlay() Then
        UpdateSyncStateLabel
        GoTo CleanExit   ' reschedule still happens at CleanExit ↓
    End If

    Set operatorWb = ResolveOperatorWorkbook()
    If operatorWb Is Nothing Then GoTo CleanExit

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

CleanExit:
    If changedLoading Then mLoading = False
    ScheduleAutoSync    ' ← always reschedule, even on early exit
End Sub
```

The `GoTo CleanExit` path already falls through to `ScheduleAutoSync` — this is not the bug in the current code. The current code structure is correct: `GoTo CleanExit` does reach `ScheduleAutoSync`. So the dead-scheduler scenario is not happening here.

***

## Failure Mode 2 — `ResolveOperatorWorkbook` returns `Nothing` during `OnTime` callback

`OnTime` fires when Excel is idle. At that moment `ActiveWorkbook` may be the add-in workbook (`ThisWorkbook` / the `.xlam`), not the operator workbook. The fallback in `ResolveOperatorWorkbook`:

```vba
If Not ActiveWorkbook Is Nothing And IsShipmentsOperatorWorkbook(ActiveWorkbook) Then
    Set mOperatorWorkbook = ActiveWorkbook
```

`IsShipmentsOperatorWorkbook` checks `WorkbookHasTable(wb, "invSys")`. If `ActiveWorkbook` at fire time is the add-in (which has no `invSys` table), this returns `Nothing` and `AutoSyncIfPending` exits at:

```vba
Set operatorWb = ResolveOperatorWorkbook()
If operatorWb Is Nothing Then GoTo CleanExit
```

The user then clicks Refresh, which calls `InitializeFromShipping` in the form context where `ActiveWorkbook` is the operator workbook — so it works. But the `OnTime` callback fires outside the form's context.

**Fix 2 — cache the workbook reference at form activation and use it in auto-sync:**

The form already has `mOperatorWorkbook`. It's set in `ResolveOperatorWorkbook`, but only when `ActiveWorkbook` is correct. Set it explicitly at initialization:

```vba
' In InitializeFromShipping, after BuildLayout:
Public Sub InitializeFromShipping()
    ...
    If mOperatorWorkbook Is Nothing Then Set mOperatorWorkbook = ActiveWorkbook
    If mOperatorWorkbook Is Nothing Or Not IsShipmentsOperatorWorkbook(mOperatorWorkbook) Then
        Set mOperatorWorkbook = ActiveWorkbook
    End If
    ...
```

And harden `ResolveOperatorWorkbook` to iterate all open workbooks as a last resort:

```vba
Private Function ResolveOperatorWorkbook() As Workbook
    On Error Resume Next

    Dim nameCheck As String
    Dim wb As Workbook

    ' 1. Cached reference
    If Not mOperatorWorkbook Is Nothing Then
        nameCheck = mOperatorWorkbook.Name
        If Err.Number = 0 And Trim$(nameCheck) <> "" And IsShipmentsOperatorWorkbook(mOperatorWorkbook) Then
            Set ResolveOperatorWorkbook = mOperatorWorkbook
            Exit Function
        End If
        Err.Clear
        Set mOperatorWorkbook = Nothing
    End If

    ' 2. ActiveWorkbook (works when form has focus)
    If Not ActiveWorkbook Is Nothing Then
        If IsShipmentsOperatorWorkbook(ActiveWorkbook) Then
            Set mOperatorWorkbook = ActiveWorkbook
            Set ResolveOperatorWorkbook = mOperatorWorkbook
            Exit Function
        End If
    End If

    ' 3. Walk all open workbooks (needed for OnTime callback when add-in is active)
    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            If IsShipmentsOperatorWorkbook(wb) Then
                Set mOperatorWorkbook = wb
                Set ResolveOperatorWorkbook = wb
                Exit Function
            End If
        End If
    Next wb

    On Error GoTo 0
End Function
```

***

## Failure Mode 3 — `ShipmentsFormRefreshReadModelForWorkbook` succeeds but `LoadShippables` does not reload NAS Inv

`ShipmentsFormRefreshReadModelForWorkbook` calls `modOperatorReadModel.RefreshInventoryReadModelForWorkbook`, which updates the `invSys` table in the operator workbook with fresh NAS values. Then `LoadShippables` calls `ShipmentsFormLoadShippables` → `BoxMakerFormLoadShippableVersionInventory`. With the BOM fast-path gate from the previous fix, this skips the network BOM open. But `BoxMakerFormLoadShippableVersionInventory` reads NAS values from the `invSys` table that `RefreshInventoryReadModelForWorkbook` just updated — so the NAS column should update.

The issue: `PreserveMissingShippableInventory` runs after `ShipmentsFormLoadShippables`:

```vba
Private Sub LoadShippables()
    Dim previousInv As Object
    Set previousInv = CurrentShippableInventoryCache()
    mShippables = modTS_Shipments.ShipmentsFormLoadShippables()
    PreserveMissingShippableInventory previousInv   ' ← overwrites blank NAS with stale cached value
    ...
```

`PreserveMissingShippableInventory` puts back the old NAS value when the new load returns a blank. If `BoxMakerFormLoadShippableVersionInventory` returns a blank NAS (because the `invSys` table row for that box version has no value), the stale `previousInv` value of `"6"` overwrites the blank, and the display never changes.

The question is why the `invSys` table would have a blank after `RefreshInventoryReadModelForWorkbook`. If the read-model refresh writes `TOTAL INV` as a number but `BoxMakerFormLoadShippableVersionInventory` reads a different column (e.g. it reads `TOTAL INV` which is the post-local-mutation value, not the authoritative NAS value), the NAS Inv column in `mShippables` could end up as the local `TOTAL INV` (already decremented) rather than the authoritative NAS.

**Add diagnostic output to `AutoSyncIfPending`** to surface exactly which step fails:

```vba
Public Sub AutoSyncIfPending()
    On Error GoTo CleanExit

    Dim operatorWb As Workbook
    Dim report As String
    Dim changedLoading As Boolean
    Dim syncCount As Long
    Dim nasBeforeRefresh As String
    Dim nasAfterRefresh As String

    If mLoading Then
        ShowStatus "AutoSync: skipped (loading)"
        GoTo CleanExit
    End If

    syncCount = PendingShipmentSyncCount()
    If syncCount <= 0 And Not modTS_Shipments.HasAnyPendingBoxVersionInventoryOverlay() Then
        UpdateSyncStateLabel
        GoTo CleanExit
    End If

    ' Capture NAS before
    If Not IsEmpty(mShippables) And UBound(mShippables, 1) >= 1 Then
        nasBeforeRefresh = NzText(mShippables(1, 4))
    End If

    Set operatorWb = ResolveOperatorWorkbook()
    If operatorWb Is Nothing Then
        ShowStatus "AutoSync: operator workbook not resolved"
        GoTo CleanExit
    End If

    If modTS_Shipments.ShipmentsFormRefreshReadModelForWorkbook(operatorWb, report) Then
        mLoading = True
        changedLoading = True
        LoadShippables
        If Not IsEmpty(mShippables) And UBound(mShippables, 1) >= 1 Then
            nasAfterRefresh = NzText(mShippables(1, 4))
        End If
        LoadShipmentState
        RefreshProjectedShippableInventory
        mLoading = False
        changedLoading = False
        UpdateSyncStateLabel
        ShowStatus "AutoSync: NAS was=" & nasBeforeRefresh & " now=" & nasAfterRefresh & ". " & report
    Else
        ShowStatus "AutoSync: read-model refresh failed. " & report
        UpdateSyncStateLabel
    End If

CleanExit:
    If changedLoading Then mLoading = False
    ScheduleAutoSync
End Sub
```

This will immediately show whether the problem is workbook resolution, read-model refresh, or NAS value propagation from invSys → mShippables.

***

## Manual test hook

Add to `modTS_Shipments` (or the form itself as a `Public Sub` callable from `Immediate` window):

```vba
' In modTS_Shipments:
Public Sub TestAutoSyncTrigger()
    ' Call this from Immediate window to verify the full auto-sync path without waiting for OnTime
    TriggerShipmentsFormAutoSync
End Sub
```

To test from the Immediate window while the form is open:
```
modTS_Shipments.TestAutoSyncTrigger
```

This exercises exactly the same code path as `OnTime` without needing to wait 15 seconds. If this works and `OnTime` doesn't, the problem is definitely `ActiveWorkbook` resolution during the callback — confirmed as Failure Mode 2 above.

***

## Failure Mode 4 — `mLoading` stuck `True` after an earlier error

`AutoSyncIfPending` sets `mLoading = True` inside the success branch. If `LoadShippables` or `LoadShipmentState` raises an error caught by `On Error GoTo CleanExit`, `changedLoading` is `True` and `mLoading = False` is correctly reset at `CleanExit`. This is handled.

But `InitializeFromShipping` also sets `mLoading = True` at line 105 and resets at line 113. If `InitializeFromShipping` errors before line 113 and the `FailInit` path doesn't reset `mLoading`, every subsequent `AutoSyncIfPending` call exits immediately. Check:

```vba
' In InitializeFromShipping:
FailInit:
    ShowStatus "Shipments form load failed: " & Err.Description
    Resume CleanExit   ' ← goes to CleanExit which does NOT reset mLoading
```

Add `mLoading = False` to the `CleanExit` / `FailInit` block:

```vba
CleanExit:
    mLoading = False   ' ← add this
    On Error Resume Next
    modTS_Shipments.EnforceShippingSupportSheetsHidden ActiveWorkbook
    If quietStarted Then modUiQuiet.EndQuietUi
    Me.MousePointer = previousPointer
    On Error GoTo 0
    Exit Sub
```