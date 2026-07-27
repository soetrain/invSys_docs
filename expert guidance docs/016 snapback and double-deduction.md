There it is. **`PickerVersionAvailableQty` calls `BoxMakerFormLoadShippableVersionInventory`, which returns `rowData(4)` (NAS) — but `ShipmentVersionInventoryBackendText` uses `rowData(4)` from the shippables array as the "backend" quantity.** Then `ApplyShipmentsSentVersionInventoryOverlay` calls `PendingBoxVersionInventoryOverlayText(rowVal, versionLabel, backendText)` where `backendText` is this picker value — and `PendingBoxVersionInventoryOverlayValue` already subtracts the pending overlay from it if one exists. So when a second Shipments Sent runs (or the same event is processed twice), the "backend" being passed in is *already* the pending-overlaid value, not the true NAS value, causing double-subtraction.

Now I have the complete picture. Let me write the full answer.

***

The report is accurate and well-written. Here is the precise technical diagnosis, ordered by root cause.

## Root Cause Stack

There are **three distinct bugs** interacting to produce the snapback and double-deduction.

***

## Bug 1 — `PickerVersionAvailableQty` feeds overlay-contaminated data back into `ApplyShipmentsSentVersionInventoryOverlay`

This is the most dangerous one.

`ShipmentVersionInventoryBackendText` calls `PickerVersionAvailableQty` to get the "authoritative backend quantity" for a version before registering the sent overlay. `PickerVersionAvailableQty` does this:

```vba
shippables = BoxMakerFormLoadShippableVersionInventory(savedBoxes)
' ...
rawQty = shippables(r, 4)  ' rowData(4) = the NAS/backend column
```

`BoxMakerFormLoadShippableVersionInventory` fills `rowData(4)` from `BuildBoxVersionInventoryCache`, which reads `tblInventoryLog`. If a prior sent overlay was persisted from a previous `Make Boxes` cycle and was **not yet evicted** (i.e., the snapshot workbook failure from the last report is still present), then `rowData(4)` already reflects that overlay delta. `ApplyShipmentsSentVersionInventoryOverlay` then receives this pre-subtracted number as `backendQty`, subtracts `qtyVal` again, and stores **that** as the new sent overlay baseline. Every form reload now displays the double-subtracted value.

**Fix:** `PickerVersionAvailableQty` must not go through the overlay path at all. It should return only the raw NAS quantity — the cached version total **before** `PendingBoxVersionInventoryOverlayValue` is applied. That means `BoxMakerFormLoadShippableVersionInventory` should separate the raw version-log total (`rowData(4)`) from the overlay-adjusted projected value (`rowData(8)`), and `PickerVersionAvailableQty` must read `rowData(4)` — which it currently does — but that requires `rowData(4)` to be guaranteed overlay-free (see Bug 2).

***

## Bug 2 — `BuildBoxVersionInventoryCache` includes `SHIP_RESERVE` and `SHIP_RELEASE` in the version aggregate

The cache accepts these five event types:

```vba
EVENT_TYPE_SHIP, EVENT_TYPE_SHIP_RESERVE, EVENT_TYPE_SHIP_RELEASE,
EVENT_TYPE_BOX_BUILD, EVENT_TYPE_BOX_UNBOX
```

A `SHIP_RESERVE` with `QtyDelta = -1` followed by a `SHIP_RELEASE` with `QtyDelta = +1` and then a final `SHIP` with `QtyDelta = -1` produces a net of `-1`, which is correct. But if the reserve or release event has a version note and the final SHIP event does **not** (or vice versa, due to the snapshot workbook failure from the prior session), the net can drift. More critically: if `SHIP_RESERVE` is present in the log but the corresponding `SHIP` has not yet been applied (NAS is mid-flight), the cache underreports NAS by the reservation amount, making `rowData(4)` look lower than it actually is on the server. The overlay then double-subtracts.

**Fix:** Version inventory NAS should only accumulate `BOX_BUILD`, `BOX_UNBOX`, and final `SHIP` rows. Reserve and release are staging artifacts, not NAS movements. Remove `SHIP_RESERVE` and `SHIP_RELEASE` from the filter in `BuildBoxVersionInventoryCache`:

```vba
' REMOVE these two from the event-type filter:
' And eventType <> EVENT_TYPE_SHIP_RESERVE _
' And eventType <> EVENT_TYPE_SHIP_RELEASE _
```

So the filter becomes:
```vba
If eventType <> EVENT_TYPE_SHIP _
   And eventType <> EVENT_TYPE_BOX_BUILD _
   And eventType <> EVENT_TYPE_BOX_UNBOX Then GoTo NextLogRow
```

***

## Bug 3 — `EvictCompletedShipmentInventoryOverlaysForShippables` uses the wrong comparison direction for a deduction overlay

The eviction condition is:

```vba
If backendQty <= overlayQty + 0.0000001 Then
    mPendingBoxVersionInventoryOverlay.Remove sentKey
```

For a SHIP event where `overlayQty = 9` (post-ship projected) and NAS catches up to `9`, this correctly evicts (`9 <= 9`). But the shippables array passed to this function comes from `BoxMakerFormLoadShippableVersionInventory`, which fills `shippables(r, 4)` via `BuildBoxVersionInventoryCache`. If Bug 2 is not yet fixed and the cache underreports, `backendQty` is lower than `overlayQty` — so the condition fires prematurely and removes the overlay before NAS has actually settled. The form then renders the raw (too-high) NAS value and the overlay is gone — the result looks like a "snapback" to the pre-ship value.

**Fix:** Once Bug 2 is fixed, the eviction direction becomes valid. But add a secondary safety check: only evict when the backend value is non-zero **and** is numeric. A blank/unknown backend must never trigger eviction.

```vba
' Before the Remove:
If Not IsNumeric(shippables(r, 4)) Then GoTo NextRow
If NzDbl(shippables(r, 4)) <= 0.0000001 Then GoTo NextRow
If backendQty <= overlayQty + 0.0000001 Then
    ' evict
End If
```

***

## The Snapback Pattern Explained

Putting it together for `T30 v2 / T31 v1 / T31 v2`:

1. `Make Boxes` runs — active pending overlay persisted from snapshot failure (Bug from prior report, partially fixed).
2. `Shipments Sent` runs for those versions.
3. `ApplyShipmentsSentVersionInventoryOverlay` calls `ShipmentVersionInventoryBackendText` → `PickerVersionAvailableQty` → `BoxMakerFormLoadShippableVersionInventory` → returns `rowData(4)` which includes SHIP_RESERVE delta from Bug 2 — returns a **lower-than-NAS** value.
4. `RegisterSentBoxVersionInventoryOverlay` stores `projectedQty = already-wrong-backend - shipped_qty`.
5. Autosync runs, `BuildBoxVersionInventoryCache` includes SHIP_RESERVE again — `rowData(4)` is still wrong.
6. Eviction fires too early (Bug 3) because `backendQty < overlayQty` — overlay removed.
7. Form re-renders with raw (inflated) `rowData(4)` — looks like it snapped back to pre-ship levels.

For `T30 v1` it happened to work because a single active version triggers the `ResolveCurrentInventoryValue` row-level fallback, bypassing the log-based cache entirely.

***

## Fix Order

| Step | File | Change |
|---|---|---|
| 1 | `modTS_Shipments.bas` `BuildBoxVersionInventoryCache` | Remove `SHIP_RESERVE` and `SHIP_RELEASE` from the event-type filter |
| 2 | `modTS_Shipments.bas` `EvictCompletedShipmentInventoryOverlaysForShippables` | Guard eviction: skip if `backendQty` is blank/zero/non-numeric |
| 3 | `modTS_Shipments.bas` `ShipmentsFormAutoSyncRefresh` | Surface `logReport` unconditionally (carried over from prior session) |
| 4 | Regression test per doc | Multi-version end-to-end test: seed → load both forms → ship 1 each → sync → assert `NAS = Projected`, no pending gauge, no overlay keys remaining |

The acceptance criteria in the doc will pass once steps 1 and 2 are applied, assuming the inventory workbook bridge resolves (the snapshot error from the prior session). The test in step 4 would have caught all three of these in advance — write that first next time.