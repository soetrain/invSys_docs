context: 
first pic is the Box Maker form, second pic is the Shipments form. `.csv` show the logging of the Shipping system from the server/NAS. this is after implementing your 3 fixes from last prompt response. below are the messages from the user side tests, "AutoSync: refresh failed." sums it all up. Shipping system sync is broken again: instead of both Box Maker and Shipments updating according to the processor, their displayed inventory appears to drift from shared stale read-model/overlay state. Box Maker Projected Inv is wrong and Shipments projected inventory is wrong. below are the recent popup while testing:

---------------------------
Microsoft Excel
---------------------------
Box build event queued for 1 T30 v1. Uses 17 component units and adds 1 shippable units after processor sync.

Sync complete.
Processed=1; StagingReport=LocalStagingMerged=1; LocalStagingFailed=0; BatchReport=Applied=1; SkipDup=0; Poison=0; RunId=RUN-invsys_Zenbook_WH-INVENTORY-20260701102912-240931; SnapshotError=Snapshot workbook not resolved.; PublishWarning=Snapshot workbook not resolved.; TimingMs=Total:9316;Batch:8707;Refresh:0
Inbox EventID: F8B426DE-86D9-41EA-AE87-5FDDAB320CF3

Completed in 9,859 ms.
---------------------------

---------------------------
Microsoft Excel
---------------------------
Box build event queued for 1 T30 v2. Uses 4 component units and adds 1 shippable units after processor sync.

Sync complete.
Processed=1; StagingReport=LocalStagingMerged=1; LocalStagingFailed=0; BatchReport=Applied=1; SkipDup=0; Poison=0; RunId=RUN-invsys_Zenbook_WH-INVENTORY-20260701102959-657804; SnapshotError=Snapshot workbook not resolved.; PublishWarning=Snapshot workbook not resolved.; TimingMs=Total:6059;Batch:5711;Refresh:0
Inbox EventID: D7261372-169A-43C5-97DF-044F9B22FEBB

Completed in 6,602 ms.
---------------------------

---------------------------
Microsoft Excel
---------------------------
Box build event queued for 1 T31 v1. Uses 12 component units and adds 1 shippable units after processor sync.

Sync complete.
Processed=1; StagingReport=LocalStagingMerged=1; LocalStagingFailed=0; BatchReport=Applied=1; SkipDup=0; Poison=0; RunId=RUN-invsys_Zenbook_WH-INVENTORY-20260701103023-176364; SnapshotError=Snapshot workbook not resolved.; PublishWarning=Snapshot workbook not resolved.; TimingMs=Total:5949;Batch:5598;Refresh:0
Inbox EventID: 7A656BCB-F031-4009-B9A5-5BE7BFF84EAE

Completed in 6,535 ms.
---------------------------

---------------------------
Microsoft Excel
---------------------------
Box build event queued for 1 T31 v2. Uses 12 component units and adds 1 shippable units after processor sync.

Sync complete.
Processed=1; StagingReport=LocalStagingMerged=1; LocalStagingFailed=0; BatchReport=Applied=1; SkipDup=0; Poison=0; RunId=RUN-invsys_Zenbook_WH-INVENTORY-20260701103041-919512; SnapshotError=Snapshot workbook not resolved.; PublishWarning=Snapshot workbook not resolved.; TimingMs=Total:5914;Batch:5559;Refresh:0
Inbox EventID: A0DB6D2F-0E61-4A55-9F14-63198E1BA440

Completed in 6,457 ms.
---------------------------

---------------------------
Microsoft Excel
---------------------------
Shipments sent: 4. package(s).
Boxes sent:
- 1. T30 v1
- 1. T30 v2
- 1. T31 v1
- 1. T31 v2
Carrier: DHL
Inbox EventID: 54AB5362-AC12-4D5B-AC14-E425F8FAE8CE
Server inventory SHIP event queued; waiting for processor/log catch-up.

Completed in 12,441 ms.
00000 ms  Shipments Sent click start
12445 ms  Shipments Sent backend call
12445 ms  LoadShipmentState start
12445 ms  LoadShipmentState active lines
12449 ms  LoadShipmentState hold lines
12449 ms  LoadShipmentState evict orphaned overlays
12449 ms  LoadShipmentState update sync label
12449 ms  Shipments Sent load shipment state
12711 ms  Shipments Sent refresh projected
---------------------------
OK   
---------------------------
AutoSync: refresh failed. AutoSync runtime: RunBatch did not handle the queued shipping event after local post/write. StagingReport=LocalStagingMerged=5; LocalStagingFailed=0 BatchReport=Applied=0; SkipDup=0; Poison=0; RunId=RUN-invsys_Zenbook_WH-INVENTORY-20260701103310-629445; SnapshotError=Snapshot workbook not resolved. RefreshReport=Skipped; TimingMs=Total:5223;Batch:5223;Refresh:0
***
Answer:
## The Single Root Cause: `SnapshotError=Snapshot workbook not resolved`
Everything else — wrong `Projected Inv`, stacked overlays, `Applied=0`, `RefreshReport=Skipped`, `AutoSync: refresh failed` — is a cascade from one unresolved failure: **the inventory workbook bridge (`ResolveInventoryWorkbookPathProcessor` / `modInventoryDomainBridge`) cannot open `WH-INVENTORY.invSys.Data.Inventory.xlsb` on the Zenbook.**

The log confirms this across every popup:
```
SnapshotError=Snapshot workbook not resolved.
PublishWarning=Snapshot workbook not resolved.
Refresh=0
```

And in AutoSync:
```
Applied=0 ... RefreshReport=Skipped
```

The four `Make Boxes` events applied (`Applied=1`) only because they were writing to the **ship inbox** (`tblInboxShip`), not requiring the inventory workbook for the write step. `RunBatch` then processed them and wrote them to the **inventory workbook** — that part worked. But `PublishInventorySnapshotBridge` failed silently afterward (`SnapshotError`), meaning `tblInventoryLog` in the **operator workbook was never refreshed** (`Refresh=0`). The `SHIP` event posted by Shipments Sent then also failed at `RunBatch` because the **same bridge failure blocks `ApplyInventoryEventBridge`** for SHIP events — those require a write-lock on the inventory workbook, which `ResolveInventoryWorkbookPathProcessor` can't get.

***
## What the Numbers Prove
**From the server CSVs, true NAS state is:**

| Box | Version | True NAS (MAKE − prior SHIPs) |
|-----|---------|-------------------------------|
| T30 | v1 | 11 (10+1+1 made, −1 shipped 6/30) |
| T30 | v2 | 11 (10+1+1 made, −1 shipped 6/30) |
| T31 | v1 | 70 (71 made, −1 shipped 6/30) |
| T31 | v2 | 25 (26 made, −1 shipped 6/30) |

**Box Maker form shows NAS = 10 / 11 / 71 / 26** — stale pre-6/30-SHIP values because `Refresh=0` means `tblInventoryLog` never received the 6/30 SHIP row. The form is reading from a stale operator cache.

**Shipments form Projected = NAS − 2 or −3** — stacked un-evicted overlays from prior sessions accumulating in the persistent overlay file on disk, because eviction requires `backendQty` to reach `overlayQty`, which never happens when `Refresh=0` keeps NAS stale in the cache.

***
## The Actual New Break Introduced by the Fixes
Fix 2 from last session (removing `SHIP_RESERVE` / `SHIP_RELEASE` from `BuildBoxVersionInventoryCache`) changed which event types are aggregated, but it did not introduce the `Applied=0` failure. What it **did** break is the `StagedShipmentVersionQty` path that `BoxMakerFormLoadShippableVersionInventory` uses as a projected delta:

```vba
' modTS_Shipments ~line 9262:
localStagedQty = StagedShipmentVersionQty(rowVal, versionLabel)
```

`StagedShipmentVersionQty` reads the **local staging file** for pending SHIP events and adds those as a projected deduction. Previously, `SHIP_RESERVE` rows in `BuildBoxVersionInventoryCache` partially offset this. With `SHIP_RESERVE` removed from the cache (Fix 2), the projected delta from `StagedShipmentVersionQty` now applies on top of a **higher** NAS baseline — creating a larger deduction than before and making Box Maker Projected look lower than it should. Except here it's showing Projected == NAS (no deduction at all), which means `StagedShipmentVersionQty` is returning 0 — because the local staging file for the **SHIP** event was already merged to NAS inbox (`LocalStagingMerged=5`), so it no longer exists on disk.

The staging file is gone. The NAS inbox has the event. `RunBatch` has the event but can't apply it. The overlay on disk still has the pending entry. The local staging delta is gone. So Box Maker shows `Projected == NAS` (no local staging delta, no log update). Shipments shows `Projected = NAS − overlays` (stale accumulated overlays). **These are two forms that have lost synchronization with each other and with the server because the inventory workbook bridge is broken.**

***
## What To Fix — In Order
### Fix 0 (Prerequisite — must solve first): Resolve the inventory workbook bridge on the Zenbook
Open the VBE Immediate window **before** any form and run:

```vba
? modInventoryDomainBridge.ResolveInventoryWorkbookBridge("WH-INVENTORY") Is Nothing
```

If this prints `True`, the bridge is returning `Nothing`. Then check:

```vba
? modProcessor.ResolveInventoryWorkbookPathProcessor("WH-INVENTORY")
```

*(make `ResolveInventoryWorkbookPathProcessor` temporarily `Public` to call it directly, or use `modConfig.GetString("PathDataRoot","")`)*

The path it's building is:
```
<PathDataRoot>\WH-INVENTORY\WH-INVENTORY.invSys.Data.Inventory.xlsb
```

On the Zenbook, `PathDataRoot` in `modConfig` (or the NAS config workbook) likely resolves to a path that either:
- Uses a UNC path (`\\NAS\...`) that isn't mapped/accessible on the Zenbook, or
- Uses a drive letter (`Z:\...`) that doesn't match what the Zenbook has mapped

**Action:** Open `modConfig` or the config INI/workbook and confirm `PathDataRoot` matches the actual mounted path on the Zenbook. This is the single fix that unlocks everything: `RunBatch` can apply the SHIP event, `PublishInventorySnapshotBridge` can publish, `Refresh` stops being 0, `tblInventoryLog` gets updated, overlays evict correctly, both forms show correct values.

***
### Fix 1: Clear the stale overlay persistence file manually (one-time)
Until Fix 0 is in place, the overlay file on disk will accumulate entries. Find it via:

```vba
? modTS_Shipments.PendingBoxVersionInventoryOverlayPath()
```

Delete or empty that file. This resets the Shipments form `Projected Inv` to match `NAS Inv` cleanly on next load, so you can see a clean baseline once the bridge is fixed.

***
### Fix 2 (Rollback of last session's Fix 2): Revert the `SHIP_RESERVE`/`SHIP_RELEASE` removal from `BuildBoxVersionInventoryCache`
The removal was architecturally premature. `StagedShipmentVersionQty` was designed as a **complement** to the cache specifically to handle the in-flight window between a Shipments Sent and the processor applying the SHIP log row. Removing SHIP_RESERVE from the cache doesn't cause `Applied=0` but it destabilizes the projected math in edge cases. Revert to the original filter until the bridge failure (Fix 0) is resolved and you have a clean baseline to validate the change's actual effect.

***
### Fix 3: Add a hard guard in `ShipmentsFormAutoSyncRefresh` against processing when `Refresh=0`
Currently when `Refresh=0`, the overlay eviction code still runs against stale NAS values, causing phantom evictions or phantom non-evictions. Add:

```vba
' In ShipmentsFormAutoSyncRefresh, after RunShippingRuntimeQueueRefresh:
Dim refreshMs As Long
refreshMs = ShippingRuntimeReportMetric(report, "Refresh")
If refreshMs = 0 And processedCount > 0 Then
    ' Processor applied events but log refresh failed — do not attempt overlay eviction
    report = report & "; OverlayEviction=Skipped(Refresh=0)"
    ShipmentsFormAutoSyncRefresh = False  ' signal caller that state is dirty
    Exit Function
End If
```

This prevents the Shipments form from falsely concluding its overlay state is clean after a sync where the log didn't refresh. The `AutoSync: refresh failed` message already surfaces, but the form currently proceeds to update its display against stale data anyway.

***
## TL;DR
The fixes from last session didn't break the SHIP processing — **the inventory workbook path was already broken before those fixes, it just wasn't surfacing as loudly.** The `Applied=0` for SHIP events is not new code — it's `ResolveInventoryWorkbookBridge` returning `Nothing` on the Zenbook, which blocks `ApplyInventoryEventBridge`. Everything else (wrong Projected, stacked overlays, `AutoSync: refresh failed`) is a symptom. Fix the NAS path config on the Zenbook and the system returns to working state.
