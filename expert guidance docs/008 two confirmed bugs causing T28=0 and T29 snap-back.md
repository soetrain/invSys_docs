Now I have the complete picture. There are two distinct bugs interacting.

***

**To Codex — two confirmed bugs causing T28=0 and T29 snap-back:**

***

## Bug A — `ShipmentsSentProjectedOverlayQty` returns `backendQty` for reserved rows, ignoring the deduction

In `ShipmentsSentProjectedOverlayQty`, the reserved-row branch is:

```vba
If isReservedRow Then
    projectedQty = backendQty               ' ← uses raw backend, NO deduction applied
    If hasExistingOverlay And backendQty <= 0.0000001 And existingProjectedQty > 0.0000001 Then projectedQty = existingProjectedQty
    If hasExistingOverlay And existingProjectedQty <= backendQty Then projectedQty = existingProjectedQty
    ShipmentsSentProjectedOverlayQty = projectedQty
    Exit Function
End If
```

When T28 is a reserved row (has `RESERVE_EVENT_ID`), `backendQty = 0` (because `ShipmentVersionInventoryBackendText` reads `invLo TOTAL INV` for ROW 89 which was corrupted to 0 by the prior hydration bug). So `projectedQty = 0`, the two overlay guards don't fire (no existing overlay yet at this point), and `RegisterSentBoxVersionInventoryOverlay` writes `SENT overlay = 0` for T28.

Then in `RefreshProjectedShippableInventory`, `activeQty = 0` (rows deleted after Sent), so the `ElseIf activeQty <= 0` branch runs:

```vba
projectedText = PendingBoxVersionInventoryOverlayText(...)   ' returns "0"
If Trim$(projectedText) <> "" And StrComp(projectedText, backendText, vbBinaryCompare) <> 0 Then
    mShippables(r, 8) = projectedText   ' "0" — displays 0
```

T28 shows `Projected Inv = 0`. The `backendText` from `mShippables(r, 4)` is also 0 (loaded from invLo which has TOTAL INV=0), so NAS Inv = 0.

For T29 (unreserved path): `projectedQty = backendQty - shippedQty = 24 - 1 = 23` → overlay written as 23. After AutoSync NAS updates to 24 (sent event processed), `PendingBoxVersionInventoryOverlayValue` checks:

```vba
If backendQty <= pendingQty + 0.0000001 Then RemovePendingBoxVersionInventoryOverlayKey sentKey
```

`backendQty=24 > pendingQty=23` → overlay not evicted. But then `EvictIdleSentOverlayForRowVersion` is called with `backendQty=24, sentProjectedQty=23`:

```vba
If backendQty <= sentProjectedQty + 0.0000001 Then   ' 24 <= 23.000 → False
```

Not evicted. But the display still shows 24 for NAS and 24 for Projected — that means `PendingBoxVersionInventoryOverlayText` is returning `backendText` (24) rather than the overlay (23). The `PendingBoxVersionInventoryOverlayValue` guard fires:

```vba
If backendQty > 0.0000001 And pendingQty <= 0.0000001 And baselineQty <= 0.0000001 Then
    RemovePendingBoxVersionInventoryOverlayKey sentKey  ' ← fires if overlay=0
```

Not that one. The actual eviction path:

```vba
If backendQty <= pendingQty + 0.0000001 _
   Or (baselineQty > 0 And backendQty < baselineQty - 0.0000001 And backendQty > pendingQty + 0.0000001) Then
    RemovePendingBoxVersionInventoryOverlayKey sentKey
```

`backendQty=24`, `pendingQty=23`, `baselineQty=24` → `24 <= 23` False. `baselineQty=24 > 0 And 24 < 24-ε` → False. Overlay survives. But T29 still shows 24 — meaning after `LoadShippables` refills `mShippables(r, 4)` with the new NAS value from the snapshot (24), the `ElseIf activeQty <= 0` branch correctly falls through:

```vba
projectedText = "23"   ' from overlay
If "23" <> "" And "23" <> "24" Then mShippables(r, 8) = "23"   ' should show 23
```

So T29 should show `Projected=23` — but the report says it jumped to 24. This means `LoadShippables` wrote `24` into `mShippables(r, 8)` (the projected column) directly, bypassing the overlay, **before** `RefreshProjectedShippableInventory` runs. That happens because `BoxMakerFormLoadShippableVersionInventory` assigns:

```vba
rowData(8) = PendingBoxVersionInventoryOverlayValue(boxRow, versionLabel, rowData(4))
```

And `PendingBoxVersionInventoryOverlayValue` is **evicting the overlay inside this call** because AutoSync just refreshed NAS to 24, and one of the eviction conditions fired. Most likely: `backendQty=24`, `baselineQty=24`, `pendingQty=23` — the condition `backendQty <= pendingQty + 0.0000001` is `24 <= 23` which is False, so that's not it. But `baselineQty` was written as `max(backendQty, projectedQty) = max(24, 23) = 24` — wait, `RegisterSentBoxVersionInventoryOverlay` sets `baselineQty = backendQty = 24` at the time of Sent. After AutoSync NAS stays 24 (the processor deducted it to 23 on the server, but the local snapshot hasn't refreshed yet — or the snapshot DID refresh and now backendQty=23 from NAS). If NAS refreshed to 23: `backendQty=23`, `pendingQty=23`, `23 <= 23` → **overlay evicted** → returns 24 (the new backend) → T29 shows 24.

That is the correct eviction — NAS caught up. **T29 showing 24 after AutoSync means NAS deducted to 23 on server, but the local read model loaded it as 24 still.** That's a snapshot-to-read-model refresh sequencing problem in AutoSync.

***

## Root cause summary and fixes:

### Fix 1 — `ShipmentsSentProjectedOverlayQty` for reserved rows must still deduct

The reserved-row branch assumes the server already holds the reservation so the deduction will happen automatically. But the **local display** still needs to show the deduction until NAS catches up:

```vba
' BEFORE (reserved row branch — returns backend, no deduction):
If isReservedRow Then
    projectedQty = backendQty
    ...
    ShipmentsSentProjectedOverlayQty = projectedQty
    Exit Function
End If

' AFTER (reserved row — deduct from backend just like unreserved):
If isReservedRow Then
    projectedQty = backendQty - shippedQty
    If projectedQty < 0 Then projectedQty = 0
    ' If existing overlay is lower (already accounts for prior deductions), honour it
    If hasExistingOverlay And existingProjectedQty < projectedQty Then projectedQty = existingProjectedQty
    ShipmentsSentProjectedOverlayQty = projectedQty
    Exit Function
End If
```

This makes reserved and unreserved rows use the same deduction formula.

### Fix 2 — `ShipmentVersionInventoryBackendText` must not return 0 from a corrupted invLo row

When `invLo TOTAL INV = 0` and `IsBlankInventoryValue(0) = False` (zero is not blank), the function returns `"0"` as the backend. Add a guard: zero is only a valid backend if it comes from the snapshot or `invSysData_Shipping`, not from a locally hydrated row with no real data:

```vba
' In ShipmentVersionInventoryBackendText, after reading invLo TOTAL INV:
If invIdx > 0 Then
    totalVal = GetInvSysValueByIndex(invLo, invIdx, "TOTAL INV")
    ' Only trust invLo if value is positive — zero may be stale from hydration
    If IsNumeric(totalVal) And CDbl(totalVal) > 0.0000001 Then
        ShipmentVersionInventoryBackendText = CStr(NzDbl(totalVal))
        Exit Function
    End If
    ' Zero or blank from invLo — fall through to snapshot/read-model
End If
```

Continue to the snapshot and `invSysData_Shipping` fallbacks as the source of truth.

### Fix 3 — `ShipmentsFormHydrateInvSysFromShippables` must never write `TOTAL INV = 0`

Already prescribed last session. Confirm this is deployed:

```vba
If totalInvQty <= 0 Then
    WriteValue lr, "TOTAL INV", vbNullString   ' leave blank, not zero
Else
    WriteValue lr, "TOTAL INV", totalInvQty
End If
```

And if the row already exists with `TOTAL INV = 0`, overwrite it with the correct value if available:

```vba
If existingIdx > 0 Then
    If NzDbl(GetInvSysValueByIndex(invLo, existingIdx, "TOTAL INV")) <= 0 And totalInvQty > 0 Then
        ' Repair the corrupted zero
        invLo.ListRows(existingIdx).Range.Cells(1, ColumnIndex(invLo, "TOTAL INV")).Value = totalInvQty
    End If
    GoTo NextShippable
End If
```

### Fix 4 — AutoSync `LoadShippables` runs before overlay eviction check

In `AutoSyncIfPending`, the sequence is:

```vba
LoadShippables              ' ← calls BoxMakerFormLoadShippableVersionInventory
                            '   which calls PendingBoxVersionInventoryOverlayValue
                            '   which may evict the SENT overlay if NAS just updated
LoadShipmentState
RefreshProjectedShippableInventory   ' ← overlay already evicted, shows NAS value
```

The overlay eviction inside `PendingBoxVersionInventoryOverlayValue` during `LoadShippables` is premature — it fires before `RefreshProjectedShippableInventory` has a chance to display the overlay value even once. The fix: **do not evict inside `PendingBoxVersionInventoryOverlayValue`**. Move all eviction to `EvictCompletedShipmentInventoryOverlaysForShippables`, which is called explicitly and only from `AutoSyncIfPending` before `RefreshProjectedShippableInventory`.

In `PendingBoxVersionInventoryOverlayValue`, remove the inline eviction calls and just return the overlay value:

```vba
' BEFORE (evicts inline):
If backendQty <= pendingQty + 0.0000001 _
   Or (baselineQty > 0 And backendQty < baselineQty ...) Then
    RemovePendingBoxVersionInventoryOverlayKey sentKey
    Exit Function   ' returns backendValue
End If
PendingBoxVersionInventoryOverlayValue = pendingQty

' AFTER (never evicts inline — just returns overlay):
' Remove ALL RemovePendingBoxVersionInventoryOverlayKey calls from this function.
' Simply return the overlay value if it exists.
PendingBoxVersionInventoryOverlayValue = pendingQty
Exit Function
```

Then in `AutoSyncIfPending`, keep the explicit eviction call before refresh:

```vba
LoadShippables
LoadShipmentState
modTS_Shipments.EvictCompletedShipmentInventoryOverlaysForShippables mShippables  ' ← evict after load
RefreshProjectedShippableInventory   ' ← now uses either evicted (NAS=authoritative) or surviving overlay
```

`EvictCompletedShipmentInventoryOverlaysForShippables` already has the correct condition (`backendQty <= overlayQty`) — it only evicts when NAS has genuinely deducted to or below the overlay value. With the eviction moved out of `PendingBoxVersionInventoryOverlayValue`, the overlay survives until the explicit eviction check confirms NAS caught up.