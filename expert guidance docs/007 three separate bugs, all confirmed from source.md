Now everything is clear. Here is the complete diagnosis:

***

**To Codex — three separate bugs, all confirmed from source:**

***
## Bug 1 — `Shipments Sent` never stages a server event for reserved rows → `Processed=0`
### Confirmed
In `ShipmentsFormRunShipmentsSentRows`, `QueueShipmentsSentEvent` is called **only for unreserved rows**:

```vba
unreservedRows = ShipmentRowsByReserveState(loShip, rowIndexes, False)
If Not IsEmpty(unreservedRows) Then
    ...
    If Not QueueShipmentsSentEvent(unreservedDeltas, ...) Then ...
End If
' ← no queue call for reserved rows
```

The normal workflow always produces **reserved rows** (because `Add` calls `QueueShipmentsReserveEvent` → stores `RESERVE_EVENT_ID`). So `ShipmentRowsByReserveState(..., False)` returns empty → `QueueShipmentsSentEvent` is never called → `StagingReport=No local staged inbox rows` → `Processed=0`.

The assumption baked into this code is: "reserved rows already have a server-side reservation so only unreserved rows need a new SHIP event." But the SHIP_RESERVE event only locked inventory on the server — it did not deduct it. The SHIP event is what tells the processor to actually deduct. It must be queued for **all rows**, reserved or not.
### Fix
In `ShipmentsFormRunShipmentsSentRows`, queue the SHIP event for **all selected rows**, not just unreserved ones. Replace the current unreserved-only queue block with:

```vba
' BEFORE (only queues for unreserved):
unreservedRows = ShipmentRowsByReserveState(loShip, rowIndexes, False)
If Not IsEmpty(unreservedRows) Then
    Set unreservedDeltas = BuildSelectedShipmentRowsDeltas(invLo, loShip, unreservedRows, "Shipments", errNotes)
    ...
    If Not QueueShipmentsSentEvent(unreservedDeltas, errNotes, queuedEventId) Then ...
End If

' AFTER (queues for all rows):
Set allSentDeltas = BuildSelectedShipmentRowsDeltas(invLo, loShip, rowIndexes, "Shipments", errNotes)
If allSentDeltas Is Nothing Then
    If errNotes = "" Then errNotes = "Unable to build shipment sent event."
    report = errNotes
    GoTo CleanExit
End If
If Not QueueShipmentsSentEvent(allSentDeltas, errNotes, queuedEventId) Then
    If errNotes = "" Then errNotes = "Unable to queue shipment sent event."
    report = errNotes
    GoTo CleanExit
End If
```

Add `Dim allSentDeltas As Collection` to the local declarations.

The existing `ApplyShipmentsSentRowsInventory` still handles the local invLo update for both reserved and unreserved rows — leave that call unchanged. Only the server-side queue call was missing for reserved rows.

***
## Bug 2 — T28 NAS Inv = 0 while T29 stays correct
### Confirmed
`ResolveCurrentInventoryValue` reads inventory from `invLo` (the local `invSys` table) **first**, before the snapshot or `invSysData_Shipping`. After `Shipments Sent`:

- `ApplyShipmentsSentDeltas` is called with `deductTotalInv = False` for **reserved rows** — it only decrements `SHIPMENTS` column, not `TOTAL INV`.
- `ApplyShipmentsSentDeltas` is called with `deductTotalInv = True` for **unreserved rows** — it decrements both `SHIPMENTS` and `TOTAL INV`.

T28 (ROW 89) was added with `Add` → reserved via `QueueShipmentsReserveEvent`. At `To Shipments` the area moves to `Shipments`. At `Shipments Sent`, `ApplyShipmentsSentRowsInventory` calls:

```vba
reservedRows = ShipmentRowsByReserveState(loShip, rowIndexes, True)   ' T28 is reserved
...
ApplyShipmentsSentDeltas(invLo, reservedDeltas, errNotes, False)   ' deductTotalInv=False
```

So T28 `TOTAL INV` in `invLo` should **not** be touched. But the screenshot shows `T28 NAS Inv = 0`.

The actual cause: `ShipmentsFormHydrateInvSysFromShippables` — added in a prior session — wrote `TOTAL INV = 0` for ROW 89 when it hydrated the row from `mShippables(r, 4)` (the displayed NAS Inv at the time, which was `0` because the BOM view had just recovered and the snapshot hadn't loaded yet, so `currentInv` was `0`). That `0` was written into `invLo`, persisted across the session, and now `ResolveCurrentInventoryValue` finds ROW 89 in `invLo` with `TOTAL INV = 0` → returns `0` → T28 shows `0`.

T29 (ROW 90) was not yet hydrated at that point (or was hydrated after the snapshot loaded with `24`), so it has the correct value.
### Fix — `ShipmentsFormHydrateInvSysFromShippables` must not write `TOTAL INV = 0`
In the hydration helper, guard against writing zero inventory:

```vba
' In ShipmentsFormHydrateInvSysFromShippables, when writing TOTAL INV:

' BEFORE:
If totalInvQty <= 0 Then
    totalInvQty = NzDbl(NzStr(shippables(r, 4)))
End If
WriteValue lr, "TOTAL INV", totalInvQty

' AFTER:
If totalInvQty <= 0 Then
    totalInvQty = NzDbl(NzStr(shippables(r, 4)))   ' displayed NAS Inv
End If
If totalInvQty <= 0 Then
    ' NAS not yet loaded — do not write 0, leave blank so ResolveCurrentInventoryValue
    ' falls through to snapshot/invSysData_Shipping
    WriteValue lr, "TOTAL INV", vbNullString
Else
    WriteValue lr, "TOTAL INV", totalInvQty
End If
```

Also: never overwrite an existing row that already has a non-zero `TOTAL INV`. The current code only writes if `FindInvRowIndexByRow(invLo, rowVal) > 0` exits early — but if a previous hydration wrote `0`, it won't be updated by a later call that has the right value. Add an update path:

```vba
Dim existingIdx As Long
existingIdx = FindInvRowIndexByRow(invLo, rowVal)
If existingIdx > 0 Then
    ' Row exists — only update TOTAL INV if current value is blank/zero and we have a real value
    Dim existingTotal As Variant
    existingTotal = GetInvSysValueByIndex(invLo, existingIdx, "TOTAL INV")
    If IsBlankInventoryValue(existingTotal) Or NzDbl(existingTotal) <= 0 Then
        If totalInvQty > 0 Then
            ' Update the existing row's TOTAL INV
            Dim existingLr As ListRow
            Set existingLr = invLo.ListRows(existingIdx)
            Dim colTotalInv As Long
            colTotalInv = ColumnIndex(invLo, "TOTAL INV")
            If colTotalInv > 0 Then existingLr.Range.Cells(1, colTotalInv).Value = totalInvQty
        End If
    End If
    GoTo NextShippable   ' already present, don't insert duplicate
End If
```

***
## Bug 3 — `SnapshotError=Snapshot workbook not resolved`
`RunShippingRuntimeQueueRefresh` → `PublishInventorySnapshotBridge` is called after `RunBatch`. Since `Processed=0` (Bug 1 not yet fixed), `PublishInventorySnapshotBridge` is currently gated off:

```vba
If processedCount > 0 Then
    If Not modInventoryDomainBridge.PublishInventorySnapshotBridge(...) Then ...
End If
```

But the status shows `SnapshotError` — meaning AutoSync is calling `PublishInventorySnapshotBridge` even at `Processed=0`. That means `ShippingRuntimeReportShowsProcessed` returned True on `Applied=0` with some other signal, or the gate condition was recently changed. Check `ShippingRuntimeReportShowsProcessed`:

```vba
If ShippingRuntimeReportMetric(batchReport, "SkipDup") > 0 Then
    ShippingRuntimeReportShowsProcessed = True
```

If `SkipDup > 0` (the event was de-duped rather than applied), `processedCount` is treated as nonzero, snapshot publish is attempted, and fails because the operator workbook can't resolve the snapshot path (`Snapshot workbook not resolved` = `PublishInventorySnapshotBridge` can't find the snapshot workbook on disk or the path config is wrong).

This is a secondary issue once Bug 1 is fixed — with real events being processed `Applied > 0`, the snapshot path will be exercised for real. The `SnapshotError` indicates the snapshot workbook path or warehouse runtime root config needs to be verified separately. That is an environment/config issue, not a code logic bug.

**Immediate action for snapshot:** confirm `target.RuntimeRoot` resolves to a writable path where `snapshot.xlsx` (or whatever the snapshot filename constant is) exists or can be created. Check `modInventoryDomainBridge.PublishInventorySnapshotBridge` for what path it constructs and verify that path on disk.