# To Perplexity / Expert - Shipments unusable performance and NAS inventory corruption

Date observed: 2026-06-25

## Executive summary

The Shipments workflow is now effectively unusable in live Excel, despite focused Phase 6 tests passing. This is no longer a single edge-case bug. It appears to be a systemic architecture failure in the Shipments form path: too much old sheet-centric/bootstrap code, read-model hydration, local `invSys` mutation, overlays, and backend repair logic are still coupled to the user-facing order-entry workflow.

The critical invariant from v4.10 is still not holding in the live system:

> `NAS Inv` must be authoritative server/read-model truth and must never be changed, reduced, inflated, repaired, or inferred from local Shipments math.

Live result contradicts that invariant:

- Shipments form load took **65,750 ms**.
- Add/order-entry control actions are also too slow.
- Sending one item took about a minute.
- After sending one item, one item inventory dropped from **13 to 0**, not 13 to 12.
- Another item changed from **19 to 24** even though it should not have been affected.
- Excel became non-responsive and visibly loaded/activated old support sheets such as `ShipmentsTally` and `InventoryManagement`.

Target expectation:

- Except for actual station-to-server processor work, all form/control operations should normally complete in **under 1000 ms**.
- Shipments form load should be at least **10x faster** than current, preferably under a few seconds on real workbooks.
- `NAS Inv` display must come only from the current server/read-model path.
- `Projected Inv` must remain a local calculator only:

```text
Projected Inv = NAS Inv - active Shipments List qty for same item/version
```

## Live symptoms

### 1. Shipments form load is unusably slow

Observed status from the live form:

```text
Loaded shipments form in 65750 ms.
```

This is after recent attempts to cache shippable version loads and stop startup sheet rebuilds. The performance problem is therefore not solved by the last incremental changes.

### 2. Control actions are also too slow

The user reports `Add` is "way too slow." Other controls in the form need complete-time instrumentation, not just load-time instrumentation.

Required instrumentation:

- form load total time;
- `LoadShippables`;
- `LoadShipmentState`;
- `RefreshProjectedShippableInventory`;
- `Add`;
- `To Shipments`;
- `Shipments Sent`;
- `Remove`;
- `Refresh`;
- auto-sync tick;
- any workbook/sheet repair/hydration call reached from those controls.

Instrumentation should show elapsed milliseconds in the form status and/or a debug log.

### 3. Old support sheets are still being loaded/activated

Excel became non-responsive while visible workbook tabs included old/support surfaces:

- `ShipmentsTally`
- `InventoryManagement`

These are not the desired user workflow surface. Their appearance suggests the form path is still calling old sheet-centric setup/repair logic or Excel is activating support sheets during table moves, formatting, or read-model hydration.

Recent attempted mitigations:

- workbook-open no longer runs full Shipments surface rebuild;
- Shipments form load no longer does BOM preflight/network refresh;
- `ShipmentsTally`, `InventoryManagement`, `ShippingBackend`, and `ShippingBOM` are explicitly hidden in some paths;
- `EnsureShippingWorkbookSurface` skips workbook-wide `AutoFit`/presentation work in the shipping bootstrap path;
- if Shipments tables are missing, open path now attempts repair instead of blocking form open.

Despite those changes, live behavior still shows old sheet loading/non-response.

## Inventory corruption observed

The live NAS/display test failed badly:

1. User sent one item.
2. That item's inventory dropped from **13 to 0**, not from 13 to 12.
3. A separate item's total changed from **19 to 24**.

This suggests at least one of the following is still happening:

- local `invSys.TOTAL INV` is being mutated by shipment paths;
- local `SHIPMENTS` / reservation / overlay math is being used to hydrate or overwrite `NAS Inv`;
- stale local read-model/snapshot values are being treated as authoritative;
- box version inventory reconstruction from logs is being used as NAS quantity;
- multiple items are being cross-keyed by package row, item name, version, SKU, or stale dictionary keys;
- old sheet repair/hydration is rebuilding `invSys` from shipment rows or BOM/version logs and then the form displays that as NAS.

## Recent relevant commits

Current implementation branch recent commits:

```text
e4858bc Repair Shipments tables during form open
41b700f Stop Shipments form startup sheet rebuilds
bb0f7e5 Cache Shipments shippable version loads
4e70ae6 Fix To Shipments local lock staging
cd5e098 Fix Shipments projected display signature
456ff15 Enforce Shipments NAS inventory boundary
```

The system still fails live after these.

## Likely high-risk code paths

### `frmShipmentsTally.InitializeFromShipping`

The form load still performs:

- carrier load;
- existing inventory preference;
- `LoadShippables`;
- `LoadShipmentState`;
- orphan overlay eviction;
- projected inventory refresh;
- support-sheet hiding.

Even with BOM preflight removed, `LoadShippables` may still reach expensive workbook/log/table paths.

### `ShipmentsFormLoadShippables` / `BoxMakerFormLoadShippableVersionInventory`

This recently gained caches, but still derives shippable rows from `BoxMakerFormLoadSavedBoxes` and version inventory from local workbook structures/logs. Need verify whether this path is still reading old support sheets or reconstructing inventory from logs in a way that contaminates `NAS Inv`.

### `GetWritableShippingInvSysTable`

This is especially suspect. It can:

- ensure/create `InventoryManagement!invSys`;
- hydrate `invSys` from shipment lines or shipping read model;
- reconcile shipment staging from shipment lines;
- reconcile shippable totals from version inventory.

This is dangerous for v4.10 if any result drives the `NAS Inv` display or mutates local `TOTAL INV`.

Any call to `GetWritableShippingInvSysTable` from form controls must be audited. It may be appropriate only for lock/floor validation and local staging, never for NAS display.

### `ReconcileShipmentStagingFromShipmentLines`

This modifies local `invSys.SHIPMENTS` based on shipment rows. That may be okay for local lock display, but it must not trigger `TOTAL INV` changes or NAS display repair.

### `ReconcileShippableTotalsFromVersionInventory`

This is a high-risk NAS boundary violation. It appears to infer or adjust `TOTAL INV` from version inventory/logs. If this runs during Shipments form open or control actions, it can explain unrelated item totals changing.

Question: should this function be removed from Shipments form paths entirely?

### Overlay functions

The v4.10 plan said sent/projected overlays must not drive `NAS Inv`. Need verify no remaining path uses:

- pending box version overlay;
- sent overlay;
- local staged version deltas;
- `invSys.TOTAL INV`;
- `Check_invSys`;
- `invSysData_Shipping`;

to set the visible `NAS Inv` column.

## Architectural concern

The current form still seems to mix four responsibilities:

1. user order-entry UI;
2. local reservation/floor guard;
3. backend/server event staging;
4. workbook/sheet/read-model repair and reconstruction.

These need separation. Shipments order entry should not be a workbook repair engine.

Recommended direction:

- Keep the form as a thin UI over an in-memory view model.
- Load `NAS Inv` from one authoritative read-model query only.
- Load active shipment rows from the local shipment list/state only.
- Compute projected locally from those two arrays only.
- Defer heavy workbook repair, BOM rebuild, log scan, read-model refresh, and support-sheet creation to explicit commands, not form open or Add.
- Never mutate `TOTAL INV` in a shipment reservation/send UI path.

## Required behavior after overhaul

### Form load

- No network refresh unless explicitly requested.
- No backend workbook open unless explicitly requested.
- No full workbook bootstrap.
- No `AutoFit`.
- No old sheet activation.
- No hidden sheet visible flicker.
- No log scanning per item.
- Target: normally under 1000-3000 ms, depending on workbook size; never 65 seconds.

### Add

- Validate requested qty immediately.
- Validation should be based on:

```text
NAS Inv - current active shipment qty - requested qty >= floor
```

- Add should create/queue the reservation/lock event.
- Add should not mutate `NAS Inv` or `TOTAL INV`.
- Add should update local list and projected display quickly.
- Target: under 1000 ms excluding actual processor/server catch-up.

### To Shipments

- Should only stage/move selected rows from Warehouse to Shipments display.
- Should be idempotent.
- Should not change NAS inventory.
- Should not run workbook repair/rebuild.

### Shipments Sent

- Should queue final shipment and clear completed local row/lock.
- Should not increase, infer, or locally repair NAS inventory.
- NAS display changes only after server processor/read-model catch-up.
- If local projection is shown during catch-up, it must remain in `Projected Inv`, not `NAS Inv`.

## Direct questions for Perplexity

1. In the current source, exactly which call chain can still cause `frmShipmentsTally` form load to take 65 seconds? Please trace from `InitializeFromShipping` down into any workbook open, sheet creation, table move, log scan, BOM rebuild, read-model refresh, overlay load, or `Application.Run`.

2. Which remaining Shipments form paths can write to `invSys.TOTAL INV` or derive a visible NAS quantity from non-NAS sources? Please list functions and call sites.

3. Should `GetWritableShippingInvSysTable` be removed from all Shipments form display paths and split into smaller functions, for example:

```text
GetNasReadModelForDisplay()
GetLocalShipmentLocksForFloorGuard()
GetMutableLocalShipmentStagingTable()
```

4. Is `ReconcileShippableTotalsFromVersionInventory` safe in any Shipments form path under v4.10, or should it be isolated to BoxMaker/manufacturing flows only?

5. What is the safest minimal architecture to make Shipments form load and Add under 1000 ms while preserving:

- NAS Inv as server/read-model truth;
- Projected Inv as local calculator only;
- Locked as local reservation/floor guard only;
- A+B event staging?

6. Should the old support sheets `ShipmentsTally` and `InventoryManagement` be retired from the user-facing Shipments form entirely, with only hidden backend tables or in-memory arrays used at runtime?

7. What exact instrumentation should be added to identify the next slow control path without adding more overhead or modal hangs?

## Proposed next-step plan

Do not continue with small overlay fixes until the call graph is cleaned up.

Recommended next implementation phase:

1. Add low-overhead timing instrumentation around every Shipments form control and load sub-step.
2. Add a debug report command that dumps timing, workbook/sheet/table access counts, and any `TOTAL INV` writes during a workflow.
3. Split NAS display loading from mutable local staging.
4. Remove or bypass `GetWritableShippingInvSysTable` from display load paths.
5. Remove or bypass `ReconcileShippableTotalsFromVersionInventory` from Shipments form paths.
6. Make form open read-only and local-first.
7. Rebuild Add/To Shipments/Sent against the clean A+B staging architecture.

## Current severity

Severity: critical / blocking.

The system is not currently usable for Shipments order entry:

- form load is about 65 seconds;
- Excel becomes non-responsive;
- Add/send actions are too slow;
- inventory display is wrong;
- NAS inventory appears corrupted or misrepresented;
- unrelated item quantities change.

