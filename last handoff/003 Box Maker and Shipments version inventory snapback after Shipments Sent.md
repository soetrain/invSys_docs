# Box Maker and Shipments version inventory snapback after Shipments Sent

Date: 2026-06-30

## Situation

After the latest Box Maker and Shipments fixes, versioned shippable inventory is visible again, but the live workflow still fails after shipping finished boxes.

Observed live state after shipping `1 ea` from each of:

- `T30 v1`
- `T30 v2`
- `T31 v1`
- `T31 v2`

The Shipments form shows:

| Box | Version | NAS Inv | Projected Inv | Locked | ROW |
| --- | --- | ---: | ---: | ---: | ---: |
| T30 | v1 | 9 | 9 | 0 | 91 |
| T30 | v2 | 10 | 10 | 0 | 91 |
| T31 | v1 | 70 | 70 | 0 | 92 |
| T31 | v2 | 25 | 25 | 0 | 92 |

Only `T30 v1` appears to reflect the shipped quantity. `T30 v2`, `T31 v1`, and `T31 v2` look like they returned to their pre-ship levels.

The Box Maker form shows the same NAS values, but Projected Inv is lower:

| Box | Version | NAS Inv | Projected Inv | ROW |
| --- | --- | ---: | ---: | ---: |
| T30 | v1 | 9 | 8 | 91 |
| T30 | v2 | 10 | 8 | 91 |
| T31 | v1 | 70 | 68 | 92 |
| T31 | v2 | 25 | 23 | 92 |

This is worse than a display-only mismatch. Box Maker is showing a likely double-deduction overlay on top of NAS quantities, while Shipments is showing some version NAS quantities that appear not to have been deducted at all.

## Why this matters

The system is supposed to preserve one boundary:

- NAS/server inventory is authoritative for stored inventory.
- Local projected inventory is temporary UI state only.
- Shipment locks/reservations are local staging state until server processing catches up.
- Versioned finished-box inventory must be read from canonical version evidence, not inferred from shared package `ROW` totals when multiple active versions exist.

The current evidence says this boundary is still not enforced across the full finished-box lifecycle.

## Current hypothesis

There are probably two interacting failures.

### Failure 1: versioned SHIP evidence is not being applied/read symmetrically with BOX_BUILD evidence

Recent fixes made `BuildBoxVersionInventoryCache` able to read canonical `tblInventoryLog` evidence for `BOX_BUILD` / made inventory. The new passing test proves this case:

- package `ROW=92`
- `T31 v1`
- canonical log row: `BOX_BUILD`, `QtyDelta=10`, note contains `ROW=92; VERSION=v1; IO=MADE`
- expected: `T31 v1` displays `10`; `T31 v2` stays blank

That test does not prove the corresponding shipment lifecycle:

- existing version inventory
- `Shipments Sent` queues/applies a `SHIP` event
- canonical inventory log records the versioned ship evidence
- `BuildBoxVersionInventoryCache` reads that negative quantity for the same version
- both Shipments and Box Maker display the deducted version quantity
- any local projected/sent overlay is cleared once NAS/log evidence catches up

The live Shipments form shows `T30 v2`, `T31 v1`, and `T31 v2` at their apparent pre-ship NAS levels. That suggests one of these is true:

- the canonical SHIP event did not include reliable version evidence for those versions,
- the processor did not write version-disambiguating SHIP log evidence,
- `BuildBoxVersionInventoryCache` is not matching SHIP rows to the finished-box SKU/version correctly,
- version-log aggregation is being overlaid or overwritten by a row-level fallback, or
- the form is reading an older/stale evidence source for some versions.

The first thing to verify is the exported canonical history for these four shipment events. It must show one distinct negative SHIP delta per version, with notes carrying enough data to distinguish `ROW=91 VERSION=v1`, `ROW=91 VERSION=v2`, `ROW=92 VERSION=v1`, and `ROW=92 VERSION=v2`.

### Failure 2: sent/projected overlays are surviving after backend catch-up and double-subtracting in Box Maker

Box Maker shows:

- `T30 v1`: NAS `9`, Projected `8`
- `T30 v2`: NAS `10`, Projected `8`
- `T31 v1`: NAS `70`, Projected `68`
- `T31 v2`: NAS `25`, Projected `23`

For a completed shipment with no active lock, once NAS/log evidence has caught up, Projected Inv should equal NAS Inv. If NAS has not caught up yet, Projected Inv should show the temporary expected post-ship value, but it must not continue subtracting after the backend already reflects the shipment.

The likely failing areas are:

- `ApplyShipmentsSentVersionInventoryOverlay`
- `RegisterSentBoxVersionInventoryOverlay`
- `PendingBoxVersionInventoryOverlayValue`
- `EvictCompletedShipmentInventoryOverlaysForShippables`
- `frmShippingBoxMaker.DisplayBoxVersionInventoryText`
- `frmShippingBoxMaker.mPendingVersionInv`

Current code has multiple overlay layers:

- persistent module overlay in `modTS_Shipments`
- form-local `mPendingVersionInv`
- sent overlay keys
- active pending overlay keys

The Box Maker screenshot strongly suggests a sent/pending overlay is still applied after NAS/log refresh. In other words, the UI is calculating:

`displayed projected = already-deducted NAS - stale local sent overlay`

instead of either:

`displayed projected = NAS` after catch-up

or:

`displayed projected = stale NAS - shipped qty` before catch-up.

## Important code observations

### `LoadShippableVersionInventoryCore`

`modTS_Shipments.LoadShippableVersionInventoryCore` now:

- reads active versions from `ShippingBOMView`,
- uses row-level `ResolveCurrentInventoryValue` only when there is exactly one active version,
- uses `BuildBoxVersionInventoryCache` for multi-version quantities,
- passes `rowData(4)` through `PendingBoxVersionInventoryOverlayValue` to produce `rowData(8)`.

This is directionally correct. The danger is that `rowData(4)` and `rowData(8)` are being treated as independent facts, when `rowData(8)` may be stale local overlay state that needs lifecycle-aware eviction.

### `BuildBoxVersionInventoryCache`

The cache now accepts events:

- `SHIP`
- `SHIP_RESERVE`
- `SHIP_RELEASE`
- `BOX_BUILD`
- `BOX_UNBOX`

and aggregates `QtyDelta` by SKU and extracted version label.

This needs a lifecycle test proving the exact signs and matching behavior. It is not enough to test one positive `BOX_BUILD` row.

Risk: if `SHIP_RESERVE` and `SHIP_RELEASE` are included in the same aggregate as `SHIP`, the read model can drift unless reserve/release deltas are known to be zero or signed exactly as intended. Versioned inventory should probably be built from inventory-affecting evidence only, not local reservation evidence, unless the contract explicitly says reserve/release are part of NAS available inventory.

### `PendingBoxVersionInventoryOverlayValue`

This function has special behavior for sent overlays:

- if sent overlay exists and backend is numeric,
- and `backendQty <= pendingQty`,
- remove sent overlay and return backend.

This condition was designed for a previous stale-baseline scenario, but it is risky because there are two different comparisons being collapsed:

- backend has caught up to the sent projected quantity,
- backend has gone below the projected quantity,
- backend is already at a post-ship value but local overlay still represents another subtraction.

The live Box Maker display suggests the overlay state is not being evicted at the right time or is being re-registered after it should have been cleared.

### `frmShippingBoxMaker.DisplayBoxVersionInventoryText`

The form-local `mPendingVersionInv` removes a pending key only when:

`backendQty = pendingQty`

That is too strict for the live lifecycle. If the server catches up to a value below the pending value, or if a separate sent overlay was registered from Shipments Sent, the form-local dictionary may continue to mask the backend value.

The form-level pending count is also only:

- `mPendingComponentInv.Count`
- `mPendingVersionInv.Count`

It does not directly reflect persistent overlay state in `modTS_Shipments`. This can make the sync gauge misleading in either direction.

## Test coverage gap

The current passing tests are too atomized. They test individual overlay math and a one-row canonical build fallback, but not the cross-form lifecycle that failed in production.

The missing regression should model the real workflow:

1. Create a shipping operator surface with two package rows:
   - `ROW 91`, `T30`, active `v1` and `v2`
   - `ROW 92`, `T31`, active `v1` and `v2`

2. Seed canonical inventory log with made/version quantities:
   - `T30 v1 = 10`
   - `T30 v2 = 10`
   - `T31 v1 = 70`
   - `T31 v2 = 25`

3. Load both forms/read paths:
   - `ShipmentsFormLoadShippables`
   - `BoxMakerFormLoadShippableVersionInventory`

4. Assert NAS and Projected are identical before any local shipment locks.

5. Stage four shipment lines, one per version:
   - `T30 v1 qty 1`
   - `T30 v2 qty 1`
   - `T31 v1 qty 1`
   - `T31 v2 qty 1`

6. Run the real Shipments Sent path.

7. Simulate processor/log catch-up by applying canonical SHIP evidence with version notes.

8. Reload both forms/read paths.

9. Assert final NAS and Projected:
   - `T30 v1 = 9`
   - `T30 v2 = 9`
   - `T31 v1 = 69`
   - `T31 v2 = 24`
   - locked count is `0`
   - pending sync count is `0`
   - no persistent overlay key remains for any of those four row/version pairs
   - no form-local pending version entry masks the backend

10. Assert no row-level fallback inflates either version from shared `ROW` total.

This one test would have caught the current live failure.

## Additional smaller tests that are still useful

### Version-log aggregate signs

Given canonical log rows:

- `BOX_BUILD +10 ROW=91 VERSION=v1`
- `SHIP -1 ROW=91 VERSION=v1`
- `BOX_BUILD +10 ROW=91 VERSION=v2`
- `SHIP -1 ROW=91 VERSION=v2`

Assert `BuildBoxVersionInventoryCache` returns:

- `v1 = 9`
- `v2 = 9`

This isolates log parsing from form overlay behavior.

### Sent overlay catch-up after backend already deducted

Register a sent overlay for:

- baseline `10`
- projected `9`

Then call display overlay with backend `9`.

Expected:

- returns `9`
- clears sent overlay
- a second call with backend `9` still returns `9`
- it must not return `8`

### Box Maker form pending dictionary catch-up

Create `mPendingVersionInv` state equivalent to:

- pending value `9`
- backend now `9`

Expected:

- display returns `9`
- pending entry is removed
- sync gauge moves toward synced

Then test:

- pending value `9`
- backend now `8`

Expected:

- display returns backend or a clearly defined stale-protection value
- it must not keep subtracting or show `7`

### Shipments form and Box Maker form parity

For the same shippable rows, both forms must render the same NAS and Projected values when there are no active local locks. The current screenshots violate this: Shipments shows projected equal to NAS, while Box Maker shows projected below NAS.

## Diagnostic data needed before code changes

Before untangling logic, capture the actual canonical log evidence for the four shipped rows.

For each of `T30 v1`, `T30 v2`, `T31 v1`, `T31 v2`, export or inspect:

- EventType
- SKU
- QtyDelta
- Note / RawNote
- ROW
- VERSION
- EventID
- source workbook/path if available

Expected evidence after shipping one each:

- one negative `SHIP` inventory event per version,
- version label present and parseable,
- no duplicate event IDs,
- no reserve/release rows being counted as final inventory movement unless intentionally signed and documented.

If the log has correct negative rows and the display is wrong, the bug is in read-model aggregation or overlay eviction.

If the log is missing negative rows or missing version labels, the bug is in Shipments Sent payload/event writing or processor application.

## Likely fix direction

Do not add another display patch first. The failure crosses Box Maker, Shipments, canonical log reads, and pending overlays.

The fix should start by defining one lifecycle invariant:

For a row/version with no active local locks:

`Projected Inv = NAS/version read-model quantity`

For a row/version with active local locks before server catch-up:

`Projected Inv = NAS/version read-model quantity - active local locks`

For a row/version whose Shipments Sent event was posted but NAS/log has not caught up:

`Projected Inv = last trusted NAS/version quantity - shipped qty`

Once NAS/log catches up:

`Projected Inv = NAS/version read-model quantity`, and all sent overlays for that row/version are gone.

The code should make this invariant explicit and testable. Right now the invariant is distributed across:

- version-log aggregation,
- `SHIPMENTS` staging,
- persistent sent overlays,
- active pending overlays,
- form-local pending dictionaries.

That distributed state is why small fixes keep passing tests but failing the live workflow.

## Acceptance criteria

The live acceptance case from the screenshots should pass:

1. Build or seed `T30 v1`, `T30 v2`, `T31 v1`, `T31 v2`.
2. Refresh Shipments and Box Maker.
3. Ship `1 ea` from each version.
4. Wait for sync/processor catch-up.
5. Refresh both forms.
6. Expected:
   - `T30 v1`: NAS `9`, Projected `9`
   - `T30 v2`: NAS `9`, Projected `9`
   - `T31 v1`: NAS `69`, Projected `69`
   - `T31 v2`: NAS `24`, Projected `24`
   - no locked rows
   - no pending inventory gauge
   - Shipments and Box Maker agree

If NAS remains `10/70/25`, debug event writing or canonical log parsing.

If NAS is correct but Projected is lower, debug overlay eviction and form-local pending state.

