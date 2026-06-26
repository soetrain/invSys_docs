Now I have the complete picture. The actual failure mechanism is now clear.

***

## Precise Diagnosis: Where ROW 89 Is Lost

The failure sequence is **not** in `QueueShipmentsReserveEvent` or `BeginShippingTableMutation`. The row is lost because **`GetWritableShippingInvSysTable` is called once and it conditionally takes the `HydrateInvSysFromShippingReadModel` path which calls `ClearListObjectData(invLo)` before writing rows.** The `invLo` VBA object reference held by `ShipmentsFormRunToShipmentsRows` still points to the same `ListObject` in Excel memory — but the `DataBodyRange` has been cleared and refilled from `invSysData_Shipping`. If ROW 89 is **absent from `invSysData_Shipping`** at that moment, `EnsureMissingInvSysRowsFromShipmentLines` adds it by calling `AddInvSysRowFromShipmentLine`. The delta is built against that freshly-added row. Then:

```
QueueShipmentsReserveEvent(deltas)
  → QueuePayloadEventCurrent
  → modRoleEventWriter.QueueEventCore
  → Opens inbox workbook (OpenWorkbookHiddenShipping / SaveWorkbookRole)
```

Opening and saving the inbox workbook triggers **Excel's automatic events pipeline** — even with `Application.EnableEvents = False` set in `BeginShippingTableMutation`. That flag is set **after** `QueueShipmentsReserveEvent` returns. Look at the order again in `ShipmentsFormRunToShipmentsRows`:

```vba
' Step 3 — Events still ENABLED here
If Not QueueShipmentsReserveEvent(deltas, errNotes, reserveEventId) Then ...

' Step 4 — Events disabled HERE, AFTER the queue
BeginShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
```

So during `QueueShipmentsReserveEvent`, `Application.EnableEvents = True` is still active. The `SaveWorkbookRole` call on the inbox workbook fires `Workbook_AfterSave` or a similar event on the **operator workbook's** auto-sync subscriber. `mShipmentsAutoSyncForm` is registered via `RegisterShipmentsFormAutoSync`. If the auto-sync fires `TriggerShipmentsFormAutoSync` → `frmShipmentsTally.OnAutoSync` → any form refresh path that calls `ShipmentsFormRefreshRuntimeInventoryCore` or `ShipmentsFormRefreshReadModelForWorkbook`, the read-model (`invSysData_Shipping`) gets **overwritten with the NAS snapshot**, and then the `invSys` table gets re-hydrated without ROW 89 (because NAS doesn't know about the pending ADD yet). The `invLo` object reference still points to the same Excel `ListObject` — it's the same pointer — but the underlying data no longer contains ROW 89.

When `ApplyShipmentDeltasLocal` then calls `FindInvListRowByRowValue(invLo, 89)`, it scans `invLo.DataBodyRange` which was just rebuilt without ROW 89, and returns `Nothing` → `"invSys ROW 89 not found."`.

***

## Three Contributing Conditions

**Condition 1 — `QueueShipmentsReserveEvent` runs before `BeginShippingTableMutation`**

`Application.EnableEvents` is still `True` when the inbox workbook save fires. Any registered worksheet change event or `Workbook_AfterSave` on the operator workbook can trigger the auto-sync path.

**Condition 2 — `invLo` is a shared mutable reference, not a snapshot**

`BuildSelectedShipmentRowsDeltas` holds `invRow` (a `ListRow`) retrieved from `invLo`. `ApplyShipmentDeltasLocal` calls `FindInvListRowByRowValue` again to re-fetch the row — it does **not** use the `ListRow` objects captured in the delta dictionaries. So even though the delta correctly carries `delta("ROW") = 89`, the re-lookup at apply-time fails when the table has been rebuilt.

**Condition 3 — `EnsureMissingInvSysRowsFromShipmentLines` repair is not durable**

The repair path in `GetWritableShippingInvSysTable` adds ROW 89 to `invSys` but does **not** write it back to `invSysData_Shipping`. So any subsequent re-hydration from the read-model (triggered by auto-sync) silently drops the repair.

***

## Fix

**Fix 1 — Move `BeginShippingTableMutation` before `QueueShipmentsReserveEvent`**

This is the immediate safeguard. Events are disabled before the inbox workbook save can trigger auto-sync:

```vba
' In ShipmentsFormRunToShipmentsRows — reorder these two blocks:
BeginShippingTableMutation loShip, previousVisibility, visibilityChanged, previousEvents, previousHandling
mutationStarted = True

If Not QueueShipmentsReserveEvent(deltas, errNotes, reserveEventId) Then ...
```

Apply the same reorder in `ShipmentsFormRunShipmentsSentRows`, `ShipmentsFormCommitLine` (ADD path), and anywhere else `QueueShipmentsReserveEvent` / `QueueShipmentsSentEvent` precedes `BeginShippingTableMutation`.

**Fix 2 — Store `ListRow` references in deltas, not just ROW numbers**

Change `ApplyShipmentDeltasLocal` to accept an optional `invRow` pre-resolved from the delta — or alternatively, make `BuildSelectedShipmentRowsDeltas` store the `ListRow` object directly in the delta dictionary (`delta("LIST_ROW") = invRow`). Then `ApplyShipmentDeltasLocal` uses `delta("LIST_ROW")` directly instead of re-calling `FindInvListRowByRowValue`. This makes the apply immune to any table rebuild between build and apply:

```vba
' In BuildSelectedShipmentRowsDeltas, inside the For Each key loop:
Set delta = CreateObject("Scripting.Dictionary")
delta("ROW") = rowKeyValue
delta("QTY") = NzDbl(requirements(key))
Set delta("LIST_ROW") = invRow    ' ← add this

' In ApplyShipmentDeltasLocal, first pass:
If delta.Exists("LIST_ROW") Then
    Set invRow = delta("LIST_ROW")
    If invRow Is Nothing Then GoTo nextDelta  ' table was rebuilt; row is gone
Else
    Set invRow = FindInvListRowByRowValue(invLo, rowVal)
End If
```

Note: VBA `Scripting.Dictionary` can hold object references — `Set delta("LIST_ROW") = invRow` works correctly.

**Fix 3 — Make the ROW 89 repair durable by writing through to `invSysData_Shipping`**

After `EnsureMissingInvSysRowsFromShipmentLines` adds a repaired row to `invSys`, mirror that row to `invSysData_Shipping` so the next hydration cycle doesn't lose it:

```vba
' After EnsureMissingInvSysRowsFromShipmentLines:
SyncRepairedRowsToShippingReadModel invLo, GetListObject(wsShip, "invSysData_Shipping")
```

This is a new small function that walks `invSys` and copies any ROW value not present in `invSysData_Shipping` into that table. This ensures the repair survives a re-hydration.