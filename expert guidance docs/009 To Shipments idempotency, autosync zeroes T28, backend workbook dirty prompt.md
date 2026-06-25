# To Codex / Expert - live Shipments failure after green tests

Date observed: 2026-06-24

## Executive summary

The current code passes the focused Shipments test slice, but the live Shipments form still fails in a real workflow. The failure now appears to be outside the narrow "Add/Shipments Sent projected overlay math" cases that tests cover.

There are three visible failures in the latest live run:

1. Clicking **To Shipments** when rows are already in the `Shipments` area is not idempotent. It clears the visible shipment rows and releases locks/projection.
2. Auto-sync then corrupts the inventory display: T28 row becomes `NAS Inv=0, Projected Inv=0`, while T29 jumps upward (`NAS Inv=24`) and remains out of sync.
3. Closing Excel prompts to save the backend inventory workbook: `invsys_Zenbook_WH.invSys.Data.Inventory.xlsb`. The user should not be exposed to this workbook or asked to save it.

The higher-level A+B plan must remain intact:

- Click path should validate auth/context and queue locally.
- Local staging should sync to NAS in a controlled path.
- Do not "fix" failures by reverting to synchronous direct NAS inbox writes from button handlers.

## Latest reproduction

Initial state in Shipments form looked mostly correct:

```text
Sync: pending (4 inventory row(s))

Box  Version  NAS Inv  Projected Inv  Locked  UOM  Location   ROW
T28  v1       15       14             1       ea   CLEARVIEW  89
T29  v1       21       20             1       EA   CLEARVIEW  90
```

Shipment display had two rows for Ref 68:

```text
Ref  Box  Qty  UOM  Area       Locked  ROW  Version  Carrier
68   T28  1    ea   Warehouse  Yes     89   v1       UPS
68   T29  1    EA   Warehouse  Yes     90   v1       UPS
```

Action 1: click **To Shipments**.

Observed after first click: looked correct.

```text
Box  Version  NAS Inv  Projected Inv  Locked
T28  v1       15       14             1
T29  v1       21       20             1

Shipment rows:
68 T28 qty 1 area Shipments locked Yes
68 T29 qty 1 area Shipments locked Yes
```

Action 2: click **To Shipments** again while rows are already in `Area=Shipments`.

Observed after second click:

```text
Box  Version  NAS Inv  Projected Inv  Locked
T28  v1       15       15             0
T29  v1       21       20             0

Shipment display is empty.
```

Expected:

- If selected rows are already in `Area=Shipments`, To Shipments should be a no-op or should only normalize carrier/area text.
- Rows must remain visible in the Shipments list.
- Locks must remain active.
- Projected Inv must remain deducted.
- It must not clear active overlays or release reservations.

Action 3: auto-sync runs after the display has already gone bad.

Observed:

```text
Sync: pending (1 inventory row(s))

Box  Version  NAS Inv  Projected Inv  Locked
T28  v1       0        0              0
T29  v1       24       20             0
```

Expected:

- T28 should never become zero from this workflow.
- T29 NAS should not jump up in a way that ignores the one reserved/shipped item.
- Auto-sync should not change server-authoritative inventory from stale local UI tables.

Action 4: close Excel.

Observed prompt:

```text
Want to save your changes to
'invsys_Zenbook_WH.invSys.Data.Inventory.xlsb'?
```

Expected:

- Backend inventory workbook should not be visible to the user.
- If it is opened transiently, it should be saved/closed intentionally by code, or opened read-only if no mutation is expected.
- The user should not be asked to decide whether to save backend authoritative data.

## Why this is not solved by the currently green tests

These tests are currently passing:

```text
20-23    role event auth/write guardrails
137-161  Shipments + refresh + BoxMaker slice
164-185  saved production, production creator, admin reconcile, role surfaces
```

The green Shipments tests cover:

- Add validation uses displayed projected inventory.
- Add does not double-deduct fresh rows.
- Update Row applies delta only.
- multi-select To Shipments keeps rows/projection for the covered setup.
- Shipments Sent does not increase projected inventory.
- SENT overlay is preserved until NAS catches up.
- BoxMaker inventory fallback is present.

The live failure appears to be a missing case:

```text
Given rows are already Area=Shipments and Locked=Yes,
When To Shipments is clicked again,
Then rows remain, locks remain, projection remains, and no overlays are cleared.
```

The current multi-select To Shipments test probably only covers Warehouse -> Shipments, not Shipments -> Shipments idempotency.

## Strong implementation suspicions

### 1. To Shipments path treats already-shipped-area rows as stale/orphaned

The second To Shipments click causes:

- shipment rows disappear,
- locks drop to zero,
- T28 projection returns from 14 to 15,
- T29 projection remains 20 but lock drops to zero.

This smells like a path that:

- filters rows by `Area="Warehouse"` and drops rows already in `Area="Shipments"`;
- calls `LoadShipmentState`, `EvictOrphanedActiveOverlays`, or `RefreshProjectedShippableInventory`;
- then sees no active shipment rows for the row/version and clears active overlays.

`To Shipments` must be idempotent for rows already in `Area=Shipments`.

### 2. Reservation identity may be lost on the second To Shipments click

If `SERVER_RESERVE_EVENT_ID` or local active-state metadata is rewritten/cleared on the second click, then `HasActiveShipmentLineForRow` may return false. That would make overlay eviction look legitimate even though the UI still represents an active shipment.

Needed checks:

- Before second To Shipments, dump `ShipmentsTally` row values:
  - `Ref`
  - `ROW`
  - `DESCRIPTION` / version
  - `Area`
  - `Locked`
  - `SERVER_RESERVE_EVENT_ID`
  - any active line id / tombstone id
- After second To Shipments, confirm whether rows were deleted from the ListObject, filtered from the form display, or moved to another table.

### 3. Auto-sync may be hydrating T28 from a damaged/stale local row

After auto-sync, T28 becomes `NAS=0, Projected=0`. We previously fixed one source of this with row-key snapshot matching and by not hydrating missing `TOTAL INV` as zero. The live failure means another path still writes or accepts zero.

Candidate paths:

- `HydrateInvSysFromShipmentLines`
- `RefreshInventoryReadModelForWorkbook`
- `ApplySnapshotToInvSys`
- `ShipmentVersionInventoryBackendText`
- any fallback that treats missing snapshot payload as numeric zero
- any path that builds `mShippables` from local `invSys` instead of the refreshed read model

Rule needed:

> Missing or unresolved backend quantity must stay blank/stale, not become zero, unless authoritative NAS snapshot says zero for the same ROW/version.

### 4. T29 NAS jump suggests stale local/NAS merge order is still wrong

T29 starts:

```text
NAS 21, Projected 20, Locked 1
```

After auto-sync:

```text
NAS 24, Projected 20, Locked 0
```

This suggests the NAS/read-model refresh is loading a later server number, but the local active reservation was already cleared by the second To Shipments click. The projection remains 20 because some overlay or stale local deduction still exists, but the lock display is gone.

This is a state-composition bug:

```text
Projected Inv = authoritative NAS/read-model base
                minus active locks
                plus/minus pending SENT overlays
```

The form is currently letting the components disagree.

### 5. Backend workbook dirty prompt means transient workbook lifecycle is broken

The prompt for `invsys_Zenbook_WH.invSys.Data.Inventory.xlsb` means the code opened the canonical inventory workbook and left it dirty/open. Possible causes:

- Auto-sync opens the inventory workbook writable, mutates read model/log/snapshot state, but does not save/close cleanly.
- Refresh path writes to canonical workbook while user-visible form still owns it.
- A helper intended to read authoritative data is using a write path.

Needed rule:

> Any canonical backend workbook opened by form refresh/sync must be either:
> - opened read-only and never mutated, or
> - explicitly saved and closed by the code path that mutates it.

The user should not be asked whether to save canonical backend state.

## Things already tried / fixed in this debugging run

### Shipments projected inventory and overlays

- Made Shipments Add fall back to displayed projected inventory when `TOTAL INV` or version ledger is stale/zero.
- Fixed Add fresh-row double deduction by registering the post-reserve `TOTAL INV` as projected overlay instead of subtracting qty a second time.
- Fixed Update Row reserved-quantity changes to apply only delta overlays.
- Fixed Shipments Sent so reserved completion keeps the projected deduction instead of adding inventory back.
- Fixed SENT overlay eviction so it clears only when NAS catches up to or below the SENT projected quantity.
- Removed/avoided sent-overlay eviction from normal form-open orphan cleanup.
- Fixed stale lower SENT overlay composition when a new active reservation is added.
- Added direct regression tests around active vs SENT overlay composition.

### To Shipments / row rendering

- Fixed the prior multi-select To Shipments case where selected rows vanished and locks cleared.
- Fixed render/list column mapping enough that messages and rows became readable.
- Current live bug is a new/narrower To Shipments idempotency case: clicking To Shipments again on rows already in `Area=Shipments`.

### Snapshot/read-model hydration

- Added ROW-key fallback when snapshot SKU differs from local display SKU:
  - example: local row `ITEM_CODE=T28`, snapshot SKU may be `T27`, but both represent `ROW=89`.
- Prevented synthetic `__ROW__` snapshot keys from creating duplicate invSys rows.
- Changed shipment-line hydration to write blank `TOTAL INV` rather than zero.

### A+B event staging architecture `0 plan docs\ln 1152-1653.md` (main content lines 1152-1653)

- Confirmed the A+B plan from docs:
  - validate auth/context on click;
  - write local durable staging;
  - sync staged rows to NAS inbox later/on demand/pre-processor;
  - dedupe by EventID;
  - do not revert button handlers to direct synchronous NAS writes.
- Updated shipping event test to accept local staging then call `SyncLocalStagedInboxRows`.
- Added production events to local staging:
  - `PROD_CONSUME`
  - `PROD_COMPLETE`
- Added receive events to local staging:
  - `RECEIVE`
- Production event creator test now passes after staging/sync alignment.

## Current known test status

Passing targeted slices:

```text
20-23    PASS  role event auth/write guardrails
137-161  PASS  Shipments + refresh + BoxMaker slice
164-185  PASS  saved production, production creator, admin reconcile, role surfaces
166      PASS  production event creator
```

Not fully green:

```text
40-44    FAIL  Warehouse bootstrap/publish cluster
older retire/archive cluster also failing in full suite
```

Those appear independent of the live Shipments bug and should not distract from the current To Shipments/autosync failure unless source confirms overlap.

## Tests needed next

### Test 1 - To Shipments idempotent for already-Shipments rows

```text
Setup:
  T28 ROW 89 v1 NAS=15, Projected=14, Locked=1
  T29 ROW 90 v1 NAS=21, Projected=20, Locked=1
  ShipmentsTally has two rows:
    Ref=68, Area=Shipments, Locked=Yes, qty=1 each
    reserve event ids present

Action:
  ShipmentsFormRunToShipmentsRows on both selected rows

Assert:
  both rows remain in ShipmentsTally
  Area remains Shipments
  Locked remains Yes
  reserve event ids remain unchanged
  T28 Projected=14, Locked=1
  T29 Projected=20, Locked=1
  no active overlay is cleared
```

### Test 2 - To Shipments second click must not clear overlays

```text
Setup:
  Same as Test 1, but inspect overlay dictionary before action.

Action:
  click To Shipments again

Assert:
  overlay keys for ACTIVE/SHIPMENTS row-version remain present
  no SENT overlay is created
  no release event is queued
```

### Test 3 - Auto-sync after To Shipments must not write zero for unresolved T28

```text
Setup:
  T28 ROW 89 has current local/NAS quantity > 0.
  Simulate snapshot/read-model row mismatch or stale missing SKU.

Action:
  Shipments auto-sync / refresh path

Assert:
  T28 NAS Inv is not overwritten to 0 unless authoritative snapshot row 89 says zero.
  If backend unresolved, display remains previous value with stale/pending status.
```

### Test 4 - Backend inventory workbook lifecycle

```text
Setup:
  Open Shipments form and trigger auto-sync/refresh.

Action:
  Close Excel/application.

Assert:
  no visible prompt to save invsys_Zenbook_WH.invSys.Data.Inventory.xlsb
  canonical inventory workbook is either closed or Saved=True
```

## Recommended next source inspection

Inspect these functions around To Shipments and refresh:

```text
ShipmentsFormRunToShipmentsRows
MarkShippingReservationRows
DeleteShipmentRows
LoadShipmentState
RenderLineList
EvictOrphanedActiveOverlays
HasActiveShipmentLineForRow
RefreshProjectedShippableInventory
ShipmentsFormRefreshRuntimeInventoryCore
HydrateInvSysFromShipmentLines
RefreshInventoryReadModelForWorkbook
ShipmentVersionInventoryBackendText
```

Specific question for source:

> On a row already in `Area=Shipments`, does `ShipmentsFormRunToShipmentsRows` treat the row as already complete/stale and clear it, or does a subsequent reload fail to include it because the loader only includes `Area=Warehouse`?

## Appendix - current live deployed code from NAS `modTS_Shipments`

Source copied from:

```text
\\100.84.136.19\invSysWH1\invSys.Shipping.xlam
LastWriteTime observed: 2026-06-24 20:32:10
Module exported: modTS_Shipments
```

Important: this is the deployed/live NAS XLAM code, not the current fork source. It is included here because the live failure was reproduced against this deployed code.

### `ShipmentsSentProjectedOverlayQty`

```vba
Private Function ShipmentsSentProjectedOverlayQty(ByVal backendQty As Double, _
                                                  ByVal existingProjectedQty As Double, _
                                                  ByVal shippedQty As Double, _
                                                  Optional ByVal hasExistingOverlay As Boolean = False, _
                                                  Optional ByVal isReservedRow As Boolean = False) As Double
    Dim projectedQty As Double

    If isReservedRow Then
        projectedQty = backendQty
        If hasExistingOverlay And backendQty <= 0.0000001 And existingProjectedQty > 0.0000001 Then projectedQty = existingProjectedQty
        If hasExistingOverlay And existingProjectedQty <= backendQty Then projectedQty = existingProjectedQty
        ShipmentsSentProjectedOverlayQty = projectedQty
        Exit Function
    End If

    projectedQty = backendQty - shippedQty
    If projectedQty < 0 Then projectedQty = 0
    If hasExistingOverlay And existingProjectedQty <= backendQty Then projectedQty = existingProjectedQty
    ShipmentsSentProjectedOverlayQty = projectedQty
End Function
```

### `ApplyShipmentsSentVersionInventoryOverlay`

```vba
Private Sub ApplyShipmentsSentVersionInventoryOverlay(ByVal invLo As ListObject, _
                                                      ByVal loShip As ListObject, _
                                                      ByVal rowIndexes As Variant)
    Dim cRow As Long
    Dim cItem As Long
    Dim cQty As Long
    Dim cDesc As Long
    Dim i As Long
    Dim rowIndex As Long
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim itemName As String
    Dim versionLabel As String
    Dim backendText As String
    Dim projectedText As String
    Dim backendQty As Double
    Dim existingProjectedQty As Double
    Dim projectedQty As Double
    Dim hasExistingOverlay As Boolean
    Dim isReservedRow As Boolean
    Dim cReserve As Long

    If loShip Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(rowIndexes) Then Exit Sub
    cRow = ColumnIndex(loShip, "ROW")
    cItem = ColumnIndex(loShip, "ITEMS")
    cQty = ColumnIndex(loShip, "QUANTITY")
    cDesc = ColumnIndex(loShip, "DESCRIPTION")
    cReserve = ColumnIndex(loShip, COL_SHIPMENT_RESERVE_EVENT_ID)
    If cRow = 0 Or cQty = 0 Or cDesc = 0 Then Exit Sub

    For i = LBound(rowIndexes) To UBound(rowIndexes)
        rowIndex = CLng(rowIndexes(i))
        If rowIndex < 1 Or rowIndex > loShip.ListRows.count Then GoTo NextRow
        rowVal = NzLng(loShip.DataBodyRange.Cells(rowIndex, cRow).value)
        qtyVal = NzDbl(loShip.DataBodyRange.Cells(rowIndex, cQty).value)
        If cItem > 0 Then itemName = Trim$(NzStr(loShip.DataBodyRange.Cells(rowIndex, cItem).value)) Else itemName = ""
        versionLabel = NormalizeBoxBomVersionLabelShipping(NzStr(loShip.DataBodyRange.Cells(rowIndex, cDesc).value))
        If rowVal <= 0 Or qtyVal <= 0 Or versionLabel = "" Then GoTo NextRow

        backendText = ShipmentVersionInventoryBackendText(invLo, rowVal, itemName, versionLabel)
        hasExistingOverlay = PendingBoxVersionInventoryOverlayExists(rowVal, versionLabel)
        projectedText = PendingBoxVersionInventoryOverlayText(rowVal, versionLabel, backendText)
        backendQty = NzDbl(backendText)
        existingProjectedQty = NzDbl(projectedText)
        isReservedRow = (cReserve > 0 And Trim$(NzStr(loShip.DataBodyRange.Cells(rowIndex, cReserve).value)) <> "")
        projectedQty = ShipmentsSentProjectedOverlayQty(backendQty, existingProjectedQty, qtyVal, hasExistingOverlay, isReservedRow)
        RegisterSentBoxVersionInventoryOverlay rowVal, versionLabel, projectedQty, backendQty
NextRow:
    Next i
End Sub
```

### `PendingBoxVersionInventoryOverlayValue`

```vba
Private Function PendingBoxVersionInventoryOverlayValue(ByVal packageRow As Long, _
                                                        ByVal versionLabel As String, _
                                                        ByVal backendValue As Variant) As Variant
    Dim key As String
    Dim sentKey As String
    Dim pendingQty As Double
    Dim backendQty As Double
    Dim baselineQty As Double

    EnsurePendingBoxVersionInventoryOverlayLoaded
    PendingBoxVersionInventoryOverlayValue = backendValue
    If mPendingBoxVersionInventoryOverlay Is Nothing Then Exit Function

    sentKey = SentPendingBoxVersionInventoryKey(packageRow, versionLabel)
    If sentKey <> "" Then
        If mPendingBoxVersionInventoryOverlay.Exists(sentKey) Then
            pendingQty = CDbl(mPendingBoxVersionInventoryOverlay(sentKey))
            baselineQty = PendingOverlayBaselineForKey(sentKey)
            If IsNumeric(backendValue) Then
                backendQty = CDbl(backendValue)
                If backendQty > 0.0000001 And pendingQty <= 0.0000001 And baselineQty <= 0.0000001 Then
                    RemovePendingBoxVersionInventoryOverlayKey sentKey
                    Exit Function
                End If
                If backendQty <= pendingQty + 0.0000001 _
                   Or (baselineQty > 0 And backendQty < baselineQty - 0.0000001 And backendQty > pendingQty + 0.0000001) Then
                    RemovePendingBoxVersionInventoryOverlayKey sentKey
                    Exit Function
                End If
            End If
            PendingBoxVersionInventoryOverlayValue = pendingQty
            Exit Function
        End If
    End If

    key = PendingBoxVersionInventoryKey(packageRow, versionLabel)
    If key = "" Then Exit Function
    If Not mPendingBoxVersionInventoryOverlay.Exists(key) Then Exit Function

    pendingQty = CDbl(mPendingBoxVersionInventoryOverlay(key))
    If IsNumeric(backendValue) Then
        backendQty = CDbl(backendValue)
        baselineQty = pendingQty
        If Not mPendingBoxVersionInventoryOverlayBaseline Is Nothing Then
            If mPendingBoxVersionInventoryOverlayBaseline.Exists(key) Then baselineQty = CDbl(mPendingBoxVersionInventoryOverlayBaseline(key))
        End If
        If backendQty > 0.0000001 And pendingQty <= 0.0000001 And baselineQty <= 0.0000001 Then
            mPendingBoxVersionInventoryOverlay.Remove key
            If Not mPendingBoxVersionInventoryOverlayBaseline Is Nothing Then
                If mPendingBoxVersionInventoryOverlayBaseline.Exists(key) Then mPendingBoxVersionInventoryOverlayBaseline.Remove key
            End If
            PersistPendingBoxVersionInventoryOverlay
            Exit Function
        End If
        If backendQty > baselineQty + 0.0000001 Then
            mPendingBoxVersionInventoryOverlay.Remove key
            If Not mPendingBoxVersionInventoryOverlayBaseline Is Nothing Then
                If mPendingBoxVersionInventoryOverlayBaseline.Exists(key) Then mPendingBoxVersionInventoryOverlayBaseline.Remove key
            End If
            PersistPendingBoxVersionInventoryOverlay
            Exit Function
        End If
    End If
    PendingBoxVersionInventoryOverlayValue = pendingQty
End Function
```

### `EvictCompletedShipmentInventoryOverlaysForShippables`

```vba
Public Sub EvictCompletedShipmentInventoryOverlaysForShippables(ByVal shippables As Variant)
    On Error GoTo CleanExit

    Dim r As Long
    Dim packageRow As Long
    Dim versionLabel As String
    Dim backendQty As Double
    Dim overlayQty As Double
    Dim sentKey As String
    Dim changed As Boolean

    EnsurePendingBoxVersionInventoryOverlayLoaded
    If mPendingBoxVersionInventoryOverlay Is Nothing Then Exit Sub
    If IsEmpty(shippables) Then Exit Sub
    If Not IsArray(shippables) Then Exit Sub

    For r = 1 To UBound(shippables, 1)
        packageRow = NzLng(shippables(r, 1))
        versionLabel = NormalizeBoxBomVersionLabelShipping(NzStr(shippables(r, 3)))
        sentKey = SentPendingBoxVersionInventoryKey(packageRow, versionLabel)
        If sentKey = "" Then GoTo NextRow
        If Not mPendingBoxVersionInventoryOverlay.Exists(sentKey) Then GoTo NextRow
        backendQty = NzDbl(shippables(r, 4))
        overlayQty = CDbl(mPendingBoxVersionInventoryOverlay(sentKey))
        If backendQty <= overlayQty + 0.0000001 Then
            mPendingBoxVersionInventoryOverlay.Remove sentKey
            If Not mPendingBoxVersionInventoryOverlayBaseline Is Nothing Then
                If mPendingBoxVersionInventoryOverlayBaseline.Exists(sentKey) Then mPendingBoxVersionInventoryOverlayBaseline.Remove sentKey
            End If
            changed = True
        End If
NextRow:
    Next r
    If changed Then PersistPendingBoxVersionInventoryOverlay

CleanExit:
End Sub
```

## Guardrails for the fix

- Do not remove local worksheet/ListObject backing tables; they are still the local form state layer.
- Do not revert A+B local staging.
- Do not make auto-sync run heavy batch work constantly on idle.
- Do not treat missing backend quantity as zero.
- Do not clear active overlays unless the corresponding local active shipment row is genuinely gone and not merely already moved to `Area=Shipments`.
- Do not expose backend canonical workbooks to the user or leave them dirty on close.
