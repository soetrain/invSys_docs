# invSys Form Controls v1

**Version:** 1.45

**Inventory date:** 2026-08-31

**Architecture:** invSys v4.11, Release 1
**Scope:** every checked-in VBA UserForm source file used by Core, Admin,
Operations, Receiving, Production, or Shipping, including active controls,
runtime-created controls, status surfaces, and generated column labels.

## 1. Purpose and authority

This is the readable Release 1 control catalog after review-checklist answers
1-3 and 6-7. It describes the implemented source and explicitly labels packaged
or visible acceptance still pending. For the deliberate D15 Production priority
change, it also records the approved target controls and separately labels the
implemented and packaged D13 acceptance state. The normative authority remains
`invSys-Design-v4.11.md`.

The catalog was derived from:

- the 17 retained `.frm` files under `src` after reviewed reachability cleanup;
- all four checked-in `.frx` designer blobs;
- the current deployed five-package set in `deploy/current`;
- runtime control-construction calls and their event handlers; and
- the 2026-08-04 and 2026-08-08 operator screenshots and acceptance reports.

The retained forms produce 17 distinct components. Item search uses the single
Core-owned `frmItemSearch` runtime canvas plus role profiles in
`Core.cDynItemSearch`; no role-named or duplicate template shell remains.

This document covers forms only. Ribbon controls and worksheet buttons are
outside its scope. The older `controls.md` remains a broader, historical
status list; this file supersedes it only for the form-control inventory.

### Status terms

| Term | Meaning |
|---|---|
| Active | A current launcher or workflow constructs and uses the form. |
| Runtime-generated | The `.frm` has an empty designer and creates controls in VBA. |
| Designer-backed | Controls are stored in a checked-in `.frx` blob. |
| Stub | A visible, intentional Release 1 placeholder, currently Purchasing in Receiving. |

## 2. Acceptance findings that affect this catalog

### Role launcher/session stability: visible checkpoint passed

Against `WHT7025AE` / `S1`, repeated Receiving, Production, and Shipping
launches reused their respective saved workbooks and forms; no additional
operator workbook opened. Activating another workbook did not rebind the forms,
and minimizing the bound workbook minimized its form with it. After restarting
Excel and repeating the launches, the session remained stable. The verified
workbook names are `WHT7025AE.Receiving.Operator.xlsm`,
`WHT7025AE.Production.Operator.xlsm`, and
`WHT7025AE.Shipping.Operator.xlsm`.

### Central Aggregator global snapshot: packaged GREEN; physical UAT pending

The advisory-only Global Inventory Snapshot now displays the exact immutable
`System_Key` beside `WarehouseId`, SKU, quantity, timestamp, and source
snapshot. It preserves distinct same-SKU entities rather than merging them by
SKU. The packaged two-isolated-warehouse proof is 10/10 GREEN, including a
same-warehouse catch-up from 8 to 11 while the other warehouse remains 5. This
does not constitute visible acceptance on two physical computers; that Release
1 checkpoint remains open.

### Demo Inventory lifecycle: packaged GREEN; seed/delete visibly accepted

`frmSeedInventory` displayed and its public callback reported one applied
processor event. The operator then refreshed, but the three demo entities were
not visible in that checkpoint. Subsequent read-only inspection found the demo
entities in all three saved role projections, and the packaged test proved the
same three unique `System_Key` values with `Condition=GOOD` through
canonical inventory, snapshot, saved Receiving workbook, and the Receiving
form's actual Refresh handler. That result is insufficient for a complete
Receiving-to-Production-to-Shipping workflow. The Demo Inventory form now
requires a data-set choice: the built-in 24-item Release 1 workflow kit or an
uploaded CSV chosen by the operator. Upload selects the source; Seed applies
it. The built-in kit contains raw ingredients, WIP, finished/shippable goods,
cases/boxes, tins, shipping cartons, dividers, labels, tape, and void fill,
with `System_Key`, `Condition=GOOD`, and catalog metadata. Repeating Seed skips
active item/location/condition groups rather than manufacturing duplicate
keys. Confirmed Delete depletes active `DEMO-` keys through audited adjustment
events without erasing history. The packaged public callback is now GREEN through the event,
canonical inventory, catalog, published snapshot, saved Receiving projection,
and the Receiving form's actual Refresh handler for the complete 24-row kit.
All 24 entities were visible, all keys were unique, every condition was `GOOD`,
and the isolated packaged Receiving -> Production -> Box Maker -> Shipping ->
restart/reconciliation chain passed 30/30. The 2026-08-09 dedicated NAS
checkpoint confirmed that the visible Seed Inventory action added inventory.
The 2026-08-19 checkpoint confirmed both Seed Demo Inventory and Delete Demo
Inventory work against the selected warehouse. The redundant bottom-right
Cancel button is removed. Closing the form without choosing an action is silent
and does not imply that a previously completed seed or delete was cancelled.

### Admin sign-in at the computer station: corrected; visible retest pending

The selected warehouse and station now resolve as `WHT7025AE` /
`X1-PRO-AI`, but the first Admin sign-in attempt found the user's capability
still scoped to the generated `S1` placeholder. The Core form-action boundary
now transitions only that successfully authenticated user's effective active
`S1` capabilities to the exact current-computer station. It preserves dates,
existing target rows, and denies, and cannot manufacture a missing role. The
focused tests are 3/3, the Core session block is 33/33, and the dedicated NAS
package run is 16/16. Visible Admin sign-in with the human credential remains
the next checkpoint.

### Receiving receipt/disposition actions: packaged GREEN; receipt retest pending

The 2026-08-19 visible checkpoint confirmed that `RETURN`, `DUMP`, and
reference concatenation worked and that disposition events appeared in Return
Entries History. The next ordinary Citric Acid receipt failed during **Add
Selected** with Excel error 1004, while a successful Confirm Dispositions
surfaced eighteen native Saving notifications. Ordinary receipt staging now
uses the same complete event-isolated tally/aggregate transaction as
disposition staging, restores the caller's prior Excel event state, and reports
the exact failed stage, error number, source, and description. The real form
confirmation handler now brackets queue, processor, refresh, and local cleanup
inside the shared quiet-UI boundary. Packaged evidence confirms quiet UI was
active during the handler and fully restored afterward; visible receipt and
Saving-notification retests remain required.

### Shipping exact-key Add and active component choices: Add visibly accepted; component retest pending

The 2026-08-20 Shipping checkpoint proved Box Maker could create shippable
boxes, but **Add** on the Shipping tab failed with `invSys table missing TOTAL
INV/SHIPMENTS/ROW columns.` The form already carried the selected immutable
`System_Key`; the local reserve apply step was the remaining legacy `ROW`
consumer. It now dispatches current-schema reservations by exact `System_Key`
and retains the numeric path only for isolated legacy test compatibility.

Box Designer's Component inventory also displayed every durable entity,
including repeated zero-balance entities. The picker now excludes nonpositive
balances and removes only repeated projections of the same `System_Key`.
Distinct positive entities remain separate even when code, item, location, and
description match. Shipping Add, Box Designer save, and Box Maker Make/Unmake
now remain inside the shared quiet-UI boundary; that boundary also hides and
restores Excel's status bar while required persistence runs. Automated
exact-key, form-action, live-role, and maintenance evidence is GREEN. The
2026-08-20 operator checkpoint confirmed Add, Update Row, Remove, Return, To
Shipments, and Shipments Sent all reached their intended actions. The filtered
component list and Saving-notification behavior still need a visible retest
against the new package.

### Shipping post-send inventory projection: visible acceptance passed

The 2026-08-20 shipment of five `12-pack Chai Box v1` packages completed and
canonical inventory correctly changed from 100 to 95 in Inventory Viewer, but
the still-open Shipping form kept displaying NAS Inv 100, Projected Inv 100,
and Locked 0. The action took about 84 seconds, including about 58 seconds in
the server processor, and surfaced twenty native Saving notices.

The transaction was correct; the stale form was a separate projection defect.
The Shipping runtime boundary ran the processor but did not refresh the
captured operator workbook afterward, and the form recalculated from its cached
shippables array. It also requested a second snapshot publication even though
`modProcessor.RunBatch` had already generated the canonical snapshot. Shipping
now reuses that processor snapshot, refreshes the captured operator read model,
reloads `lstShippables`, and only then derives Projected Inv and Locked. The
duplicate publication and one duplicate legacy stage-cleanup call are removed;
the processor's required three durability saves remain.

Focused post-send contracts are 4/4, packaged XLAM is 74/74, live role workflows
are 47/47, and the ordered Release 1 chain is 30/30. In the live public form
action, shipping five units from twenty refreshed the same operator workbook to
fifteen before the form action returned. The 2026-08-20 visible retest then
shipped six packages and confirmed that the still-open Shipping list showed the
correct NAS Inv, Projected Inv, and Locked values. The Viewer agreed with the
deducted balance. Total time improved from about 84 seconds to about 65 seconds;
the report attributed about 49 seconds to processor batch persistence and 836
milliseconds to canonical read-model refresh.

Excel still surfaced native Saving progress notices during required NAS writes:
four during a one-row **Add**, with additional notices during Shipments Sent.
Slice 4u therefore preserves the inbox, reservation-ledger, and processor
durability boundaries, adds one `Persistence summary:` line to the existing
Shipping status/message output, and batches all reservation-ledger row updates
into one workbook save per multi-row action. Native Excel save-progress windows
cannot be reparented into a VBA form; the invSys summary is the authoritative
single action report, while Excel may still display its own progress UI for the
two logical server saves required by a one-row Add.

### Receiving and Production persistence feedback: packaged GREEN; visible retest pending

Slice 4v extends the same feedback contract to the other operator forms.
Successful **Confirm Writes** and **Confirm Dispositions** append one
`Persistence summary:` line to Receiving's existing multiline status box after
the batched inbox write, processor application, read-model refresh, and staging
cleanup. Successful Production **Complete Run** appends one summary to
Production's existing multiline status box after its consume/complete events
and processor work. Both public action paths use the shared quiet-UI boundary.

The shared Receiving/Production run-and-refresh path no longer publishes a
second snapshot after `modProcessor.RunBatch`; the processor remains the sole
snapshot and durability owner. Its required inventory, outbox, and inbox saves
are unchanged. Native Excel save-progress windows cannot be moved into a form,
so any that remain must be counted separately from the one invSys status
summary. Focused persistence feedback is 4/4, packaged XLAM is 74/74, live role
workflows are 47/47, the ordered Release 1 chain is 30/30, and reviewed cleanup
is 13/13.

### Reusable Production Processes and Recipe graphs: packaged workflow GREEN

Architecture v4.11 D15 and Plan 022 Slice 4x deliberately replace the current
single **Recipe Builder** page with **Process Designer** and **Recipe Designer**.
The packaged form now has five top-level pages: Process Designer, Recipe
Designer, Ingredients Assignment, Production Run - List, and experimental
Production Run - Tree. The former four-page Recipe Builder surface is removed.
The List-run now loads released reusable Recipes and executes their Process
graph; the legacy single-builder path is not a silent fallback.

Process Designer owns named Process draft/save/release/obsolete/reuse actions,
requirements, acceptable SKU alternatives, instructions, and one or more output
definitions. Recipe Designer selects exact released Process versions, connects
individual outputs to downstream requirements, controls execution order, and
shows unresolved-input, compatibility, quantity, and circular-dependency
validation. Production Run - List preserves `0.001%` through `1000%` scaling,
allocates exact available inventory `System_Key` entities, creates each Process
output under a distinct new key, consumes routed intermediate outputs by that
same key, and leaves unconnected output balances as finished/co-product
inventory. The existing launcher, captured saved workbook, persistence summary,
quiet-UI boundary, and experimental Tree scope remain unchanged.

The headless Designs Domain boundary is implemented and automated GREEN for
Process/Recipe schema, lifecycle events, immutable versions, graph validation,
released-only lists, exact-version serialized definitions, and released-Recipe
validation. The same read contracts are GREEN through the Core cross-XLAM and
Operations primitive bridges. The reusable Production range is 17/17 GREEN,
packaged five-XLAM validation is 74/74 GREEN, and packaged Ribbon validation is
142/142 GREEN. The final static baseline contains 152 components, 4,987
procedures, 1,033 candidates, 8 literal `Application.Run` targets, 45 unresolved
dynamic calls, and 189 duplicate-body groups; deliberate Slice 4x growth is
bounded by the 13/13 reviewed exception contract.

The packaged public-launcher form-action path is GREEN for Process
save/release/obsolete/reuse, assignment-backed version creation, and Recipe
select/connect/order/save/release/obsolete. The same run proves captured
saved-workbook binding and second-launch reuse. Minimum/default/expanded and
native maximize/restore geometry is GREEN across all five pages. The actual
List-run handlers are GREEN for inclusive scaling, exact-key allocation,
insufficiency/stale rejection, two multi-output batches, routed intermediate
consumption, a 20%-of-10-unit co-product yield basis, retained co-products, and
correlated persistence. Public Viewer
Events exposes **Production Input Consumed** and **Production Output Created**.
Clean-session saved-workbook restart is GREEN through the public launcher and
actual Run List **Load** handler: a second Excel process reopens the same saved
Production workbook and loads the persisted exact released Recipe. Dedicated-
NAS launcher readiness is 16/16 across two clean sessions with no canonical
file changes from launcher use. Visible reusable-workflow and five-page layout
UAT remain open.

### Slice 4y Process worksheet editing: packaged GREEN, visible UAT pending

The 2026-08-23 checkpoint accepted **Upload Data Set** and **Seed Demo
Inventory**. The supplied screenshot is Inventory Viewer and proves that the
uploaded seeded sample is visible there, so Viewer remains accepted and
unchanged.

The same checkpoint rejected manual/GUID Process identities and the unexplained
display word **Basis**. The approved target uses locked, automatically generated
three-character Base-36 Process, Recipe, Requirement, and Output IDs. Raw
**Basis** wording is replaced by **Batch basis quantity** for percentage inputs
and **Yield basis quantity** for percentage outputs. Slice 4y first proved a
single-table round trip at focused source 6/6 and packaged callback/restart 2/2;
its toggle was then deliberately superseded by Slice 4z. The dedicated NAS
launcher gate remains 16/16 GREEN with no canonical file changes. Visible
formulation-table and five-page layout UAT remain pending.

### Slice 4z multi-table import workbench: packaged GREEN; visible UAT pending

The 2026-08-24 checkpoint visibly accepted the first generated Process table
on **invSys Process Editor**. It rejected the single toggle/single-table model,
operator-authored Percent, output Item Code and Design ID, free-text Record Type,
and omission of Ingredient Assignment data. The approved target has separate
create/retrieve actions, multiple simultaneous selected tables, calculated
Percent/basis columns, validation dropdowns, generated output Design identity,
and item-search-backed acceptable managed items. Packaged RED/GREEN is complete;
visible retest remains pending under Plan 022 Slice 4z.

The focused source contract is now 7/7 GREEN and the packaged public-launcher
gate is 2/2 GREEN. The packaged run created three simultaneous Process tables,
retrieved only the selected first table, preserved the other two through a clean
Excel restart, and retrieved those remaining tables one at a time. Formula
evidence is 611.2 lb with 16.4/32.7/1.8/49.1 percentages totaling 100.0%.
Final regression evidence is XLAM 74/74, Ribbon/compile 142/142, live roles
47/47, ordered Release 1 chain 30/30, launcher contracts 24/24, dedicated NAS
16/16, deterministic static baseline 19/19, and reviewed cleanup 13/13. The
initial expanded-page layout RED exposed a Retrieve-button/Description overlap;
the relocated 170-point **Retrieve Selected Process** control is GREEN across
all five pages and native minimize/restore/maximize transitions. Visible
operator workbench UAT remains pending.

### Slice 4aa text-safe IDs, catalog UOMs, and bulk import: packaged GREEN; visible UAT pending

The accepted worksheet Percent formulas remain unchanged. The Process table now
requires text-safe generated Base-36 metadata/row identities, automatic INPUT
Requirement IDs, UOM validation from Settings' **Recipe UOM Catalog**, numbered
acceptable managed-item/SKU pairs on INPUT rows, actual Core picker opening on
cell entry, and Ctrl+click multi-table DRAFT import. This is a deliberate D15
refinement protected through the public Production launcher, worksheet event,
and Process create/retrieve handlers.

The 2026-08-24 focused RED was `1/8`, with seven missing behavior contracts.
Focused source is now `8/8` GREEN and historical Slice 4z source remains `7/7`
GREEN. The packaged Production/restart proof is `2/2` GREEN and records all
eight bulk-import fields true, including actual `PickerOpened`, multi-area
selection, and multi-table DRAFT persistence through the public form handler.
The full packaged, role-workflow, Release 1 chain, NAS, static, reviewed-growth,
and five-page layout regressions are GREEN. Operator-visible UAT remains
required before acceptance is checked complete.

### Slice 4ab Process picker inventory projection: inventory visibility accepted; selection UAT pending

The 2026-08-24 operator checkpoint accepted Process-table creation, multiple
tables, and automatic opening of **Production Item Search**. The opened picker
showed no rows even though seeded managed inventory was visible in Inventory
Viewer. This is a blocker against the existing control contract: every
acceptable-item picker result must come from the current warehouse managed
inventory projection, use exact nonblank `System_Key` identity internally, and
write only the selected managed item/SKU alternative to the Process table. The
picker must not require or expose legacy `ROW`, and selecting an alternative
must not allocate an inventory entity.

The focused source contract is `4/4` GREEN. The packaged public Production
callback and clean-restart proof is `2/2` GREEN and records both
`PickerOpened=True` and `PickerInventoryRows=True` after resolving active exact
Inventory Domain entities and deduplicating them to acceptable SKU alternatives.
Packaged XLAM, Ribbon/compile, live role, ordered Release 1 chain, launcher, NAS,
static, and reviewed-growth regressions are GREEN. On 2026-08-25 the user
visibly confirmed that seeded warehouse inventory now appears; selecting an
item into the expected numbered acceptable managed item/SKU pair remains in the
final workflow acceptance.

### Slice 4ac Process OUTPUT managed-item picker: packaged proof superseded by Slice 4ad RED

The 2026-08-25 operator checkpoint accepted that Production Item Search now
contains current managed inventory and identified the missing equivalent action
for OUTPUT rows. Under the revised D15 contract, entering an OUTPUT **Name**
cell opens that same picker. Commit writes the visible managed item name and a
hidden, system-managed **Output SKU**; no Item Code or `System_Key` is authored
by the operator. Output and Design IDs remain generated, and Production run
completion remains the only point that creates a physical output under a new
`System_Key`.

Focused RED was `1/6`; focused source is now `6/6` GREEN. The packaged public
Production callback plus clean restart is `2/2` GREEN and proves the actual
worksheet selection event and Core picker commit with all five OUTPUT contract
fields true. Packaged XLAM, Ribbon/compile, live role, Release 1 chain,
launcher, dedicated NAS, static, and reviewed-growth gates are GREEN. Visible
confirmation that OUTPUT search, selection, and retrieval behave correctly in
the user's saved Production workbook remains pending.

The subsequent visible checkpoint showed the packaged test targeted OUTPUT
Name while the operator used the same visible **Acceptable Managed Item 1**
column that opens search for INPUT. That OUTPUT path remained unreachable, so
Slice 4ac's automated result is retained as historical evidence but does not
satisfy visible acceptance.

### Slice 4ad Process picker INPUT/OUTPUT reachability: packaged GREEN; visible retest pending

Production Item Search must open when the operator enters **Acceptable Managed
Item 1** for either INPUT or OUTPUT. INPUT commit fills its numbered alternative
pair. OUTPUT commit fills the visible managed-item selector and hidden **Output
SKU**, while preserving a nonblank descriptive Output Name and never retaining
a physical `System_Key`. OUTPUT Name may remain a convenience picker target but
cannot be the only OUTPUT route.

Focused source RED was `5/6`; the corrected packaged public-handler RED was
`0/2` and recorded the three OUTPUT picker fields false while adjacent
Production workflows remained GREEN. Focused source is now `6/6` and packaged
Production plus clean restart is `2/2`, with open, commit, SKU round-trip,
descriptive-name retention, and no-physical-key fields all true. Packaged XLAM,
Ribbon/compile, live role, Release 1 chain, launcher, dedicated NAS, static, and
reviewed-growth gates are GREEN. The user's visible INPUT/OUTPUT retest remains
required.

### Slice 4ae Process picker cell boundary and row identity: packaged GREEN; visible retest pending

Production Item Search opens only from **Acceptable Managed Item n** cells:
all numbered alternatives on INPUT/REQUIREMENT and **Acceptable Managed Item
1** on OUTPUT. OUTPUT Name and every other Process-table cell are not picker
targets. INPUT, REQUIREMENT, OUTPUT, and INSTRUCTION rows share one automatic,
text-safe, three-character Base-36 row-ID namespace across the whole table;
`000` is reserved, existing unique IDs remain stable as rows are entered in any
order, and duplicates are corrected by the public worksheet change handler.

Focused source RED was `0/6`; packaged ProductionReusable RED was `1/2` with
OUTPUT Name suppression and row uniqueness false while the prior picker path
remained functional. Focused source is now `6/6`; packaged ProductionReusable
is `2/2` and records `OutputNamePickerSuppressed=True`, `UniqueRowIds=True`, and
`FirstAssignedIdRetained=True`. Packaged XLAM, Ribbon/compile, live role,
Release 1 chain, launcher, dedicated NAS, deterministic static, and reviewed
growth gates are GREEN. Rebuilt packages are in `deploy/current`; visible
operator acceptance remains pending.

### Slice 4af Recipe identity initialization and editable version: packaged GREEN; visible retest pending

Recipe Designer initializes a blank draft with the next collision-checked
three-character Base-36 Recipe ID and proposed version `1` after the form loads,
on **New Recipe**, and on **Clear**. Save Draft and Release also supply either
missing generated value before validation. Recipe ID remains locked. Recipe
Version is editable, accepts a positive whole number, and preserves the
operator's valid nonblank value through lifecycle handlers; immutable existing
versions are never overwritten. Focused test-first RED was `0/6`; focused
source is `6/6` GREEN and packaged Production plus clean restart is `2/2`
GREEN through the actual initialization, New Recipe, Save Draft, and Release
paths. The packaged action changes Version from `1` to `9` and records the
edited value retained. Rebuilt packages are registered; visible acceptance is
pending.

### Slice 4ag reusable-run actual output and list readability: visibly accepted

Production Run - List uses the scaled definition as **Planned**, but every
reusable output requires its own positive **Actual Output** before completion.
The actual quantity governs created managed inventory and becomes **Last
Actual** after completion; planned quantity remains visible. Palette, Inventory
Check, and Production Output display exact identity as readable **System_Key**,
not **Inventory ID**. Focused test-first RED was `0/6`; focused source is `6/6`
GREEN and packaged Production plus clean restart is `2/2` GREEN through the
actual output-row selection/change and Complete Run handlers. Packaged evidence
records `ActualOutputAccepted=True`, `LastActualDisplayed=True`,
`ActualInventoryQty=True`, and `SystemKeyHeadersReadable=True`. Full preserved
regressions are GREEN. Visible acceptance on 2026-08-26 confirmed the completed
tea batch's Last Actual, newly managed output inventory, and reusable batch
behavior.

### Slice 4ah Admin inventory edit selection binding: visibly accepted

In Edit Item mode, selecting a visible `cmbEditItem` dropdown row must bind that
row's exact catalog SKU and load its editable fields before Save. Typed search
may filter candidates but is not itself a selection. Combo-dropdown and
`lstEditItemResults` choices must use the same SKU-backed loader. Selecting
**Utility** in Qty mode for the bound item emits `TRACK_QTY=FALSE` and
`ITEM_KIND=UTILITY`, so a numeric tank quantity is not required. Focused
test-first RED was `0/5`; focused source is `5/5` GREEN and packaged XLAM/Admin
handler validation is `75/75` GREEN. The packaged handler records
`ComboSelected=True`, `FieldsLoaded=True`, `UtilityReady=True`, and
`ValidationReady=True`. Visible acceptance on 2026-08-26 saved Filtered Water as
Utility without the false selection error. Slice 4aj separately corrects the
Production quantity display for that accepted catalog state.

### Slice 4ai Admin inventory worksheet workbench: packaged GREEN

The Admin ribbon launcher wording changes to **Add/Edit Inventory Items** while
retaining `btnAdminAddInventoryItem`, `modAdmin.Add_InventoryItem`, and
`ADMIN_MAINT`. `frmAddInventoryItem` adds **Create Inventory Table** and
**Upload Selected Inventory Table**. The first creates a uniquely named local
staging table on **invSys Inventory Editor** in the captured workbook and then
returns control to the worksheet; the second resolves the selected invSys table
in that same workbook and runs whole-table preflight before existing Admin
catalog/event/processor actions. Multiple tables and pasted CSV-shaped rows are
allowed. ADD Item Code is generated; EDIT uses an exact existing Item Code and
requires a reason. UOM and Qty Mode are dropdown-backed, extra columns become
custom fields, `ROW`/`System_Key` are prohibited, and Upload Status/Result make
successes and retryable failures visible. The focused contract is `8/8` GREEN,
packaged action validation is `76/76`, and Ribbon/compile validation is
`142/142`. The packaged action invokes the real form handlers in a captured
saved workbook and records table creation, whole-table preflight, counted and
Utility ADD, exact EDIT, generated code, and row-status evidence. Release 1,
launcher, NAS, static, and growth regressions are GREEN; visible acceptance is
pending.

### Slice 4aj Production batch history and Utility projection: packaged GREEN

Production Output retains a distinct completed row per batch and Process output
instead of overwriting the prior batch. The active batch remains selectable for
the existing Actual Output handler. The visible heading **Planned** becomes
**Used Goods**, showing scaled compatible inputs consumed by the Process, and a
new **Process Total** column shows cumulative actual output for the same
Process/output/UOM across retained rows. Each completed row keeps its own Last
Actual, Batch, Recall, and exact `System_Key`. Run palette **Inv** and Inventory
Check display **Utility** when catalog metadata says `TRACK_QTY=FALSE` or
`ITEM_KIND=UTILITY`, while required/used quantities continue to show measured
usage. Focused RED was `1/7`; focused GREEN is `7/7`. Packaged Production plus
clean restart is `2/2` and records `BatchHistoryRows=True`,
`ProcessTotal=True`, and `UtilityDisplay=True` through the real operator
handlers. XLAM/Ribbon, live-role, Release 1, NAS, deterministic-static, and
reviewed-growth gates are GREEN. Visible acceptance is pending.

### Slice 4ak Admin-created inventory visibility and dropdowns: automated GREEN, visible acceptance pending

**Add Item** and inventory worksheet `ADD` create a managed entity through
`INVENTORY_CREATE` with a new `System_Key`; the operator path may not write a
blank-key `MIGRATION_SEED` quantity. After processor/snapshot Refresh, a counted
positive-quantity item is visible in Inventory Viewer and the Production
managed-item picker. An explicit Edit/Save can create the first entity for a
catalog-only item produced by the superseded Add defect when the operator
supplies a positive target. **Default location** and **Category** are editable
dropdowns populated from distinct current warehouse catalog values, with the
configured default included for location. Focused RED was `2/7`; focused GREEN
is `7/7`. The packaged Admin test invokes the real **Add Item** handler and
records `SubmitHandler=True`, `ExactEntityCreate=True`,
`LocationDropdown=True`, and `CategoryDropdown=True`; the full packaged handler
gate is `77/77`. Preserved evidence is: launcher source `24/24`,
ProductionReusable packaged and restart `2/2`, Ribbon/compile `142/142`, live
role workflows `47/47`, ordered Release 1 `30/30`, dedicated NAS runtime
`16/16`, deterministic static `19/19`, and reviewed growth/cleanup `13/13`.
Visible acceptance remains open for Honey in Viewer and the Production picker
and for both dropdowns in the deployed form.

### Slice 4al Add/Edit managed-inventory deletion: automated GREEN, visible acceptance pending

In Edit Item mode, **Delete Item** becomes enabled only after the operator
selects an exact catalog-backed SKU. The real button handler requires a second
confirmation and a nonblank reason. It retires every active exact `System_Key`
entity for that SKU through one audited `ADMIN_INVENTORY_ADJUST` event: counted
balances are reduced to zero and Utility/Service entities receive zero-delta
retirement evidence. Catalog, entity, and event history remain durable; after
Refresh, retired rows are omitted from ordinary Edit search, managed inventory,
Inventory Viewer levels, and Production pickers. Focused RED was `1/7`. A
post-deployment declaration-scope regression then recorded `7/8` RED before the
real public handler declarations were corrected; focused GREEN is now `8/8`.
Packaged action validation is `79/79`. Ribbon/compile
`142/142`, live roles `47/47`, ordered Release 1 `30/30`, dedicated NAS `16/16`,
deterministic static `19/19`, and reviewed growth/cleanup `13/13` remain GREEN.
Visible acceptance is pending.

### Slice 4am zero starting quantity: automated GREEN, visible acceptance pending

Counted **Add Item** and inventory worksheet `ADD` accept an explicit numeric
Starting Qty of zero and reject negative values. Zero creates one active exact
`System_Key` entity and remains visible in managed inventory, Inventory Viewer,
Receiving choices, and Process Designer's managed-item picker after Refresh.
It is not allocatable stock: Production Run exact-key allocation still lists
positive available counted entities only. Retirement remains the explicit
Delete Item action. Focused RED was `1/8`, with a strengthened rebuild-state RED
at `7/8`; focused GREEN is `8/8`. Packaged action validation is `81/81`,
Ribbon/compile is `142/142`, live roles are `47/47`, ordered Release 1 is
`30/30`, dedicated NAS runtime is `16/16`, deterministic static is `19/19`,
and reviewed growth/cleanup is `13/13`. Visible acceptance is pending.

### Slice 4an mixed-UOM Process assembly: automated GREEN, visible acceptance pending

Process worksheet INPUT **Basis Qty** and **Percent** are grouped by normalized
UOM. Each group has its own basis subtotal and percentages totaling 100.0%; LB
material and EA packaging may therefore coexist in one Process and scale from
their declared quantities. **Retrieve Selected Process** must preserve every
row's UOM and must not invent an LB-to-EA conversion. Blank or catalog-invalid
UOM, invalid quantities, and a UOM group's invalid percentage total remain
errors. Recipe connections still require compatible UOM. Packaged public-
handler RED was `0/2`; focused GREEN is `7/7` and packaged Production plus clean
restart is `2/2`, with `MixedUomAccepted=True` and
`MixedUomRowsPreserved=True`. Packaged XLAM `81/81`, Ribbon/compile `142/142`,
live roles `47/47`, ordered Release 1 `30/30`, dedicated NAS `16/16`,
deterministic static `19/19`, and reviewed growth/cleanup `13/13` are GREEN.
Visible retrieval acceptance is pending.

### Slice 4ao Recipe connection labels and UOM control: automated GREEN, visible acceptance pending

Recipe Designer's connection Output control displays the selected Process
output name while retaining the generated Output ID as its hidden bound value.
The connection UOM control is `cmbConnectionUom`, a dropdown-list populated
from Settings' current Recipe UOM Catalog. Connect and Update continue to stage
the hidden Output ID and the selected catalog UOM. Focused RED was `0/7` and
GREEN is `7/7`; packaged public-handler plus clean-restart evidence is `2/2`
with name display, hidden ID persistence, catalog selection, and Update all
true. Packaged XLAM `81/81`, Ribbon/compile `142/142`, live roles `47/47`,
Release 1 `30/30`, NAS `16/16`, static `19/19`, and growth `13/13` remain
GREEN. Visible acceptance is pending.

### Slice 4ap named Recipe graph layout: packaged GREEN, visible acceptance pending

Recipe Designer displays Process, Output, and Requirement names while retaining
generated IDs as hidden selection/persistence keys. Released Processes and
Recipe Process Nodes appear above a full-width Connections projection; both
node and connection projections have aligned column headers. The connection
editor explicitly labels upstream and downstream Processes, rejects a
same-node self-reference, and states that a final/co-product output is
intentionally left unconnected and becomes managed inventory. Focused
`2/9` RED and `9/9` GREEN are recorded. Packaged Production/restart is `2/2`,
including name display with hidden identity, aligned headers, full-width
geometry, visible selection, Connect/Update/Disconnect, same-node rejection,
and finished-output guidance. Packaged XLAM `81/81`, Ribbon/compile `142/142`,
live roles `47/47`, ordered Release 1 `30/30`, NAS `16/16`, static `19/19`, and
growth `13/13` remain GREEN. Visible acceptance is pending.

### Slice 4aq output-first Recipe routing: packaged GREEN, visible acceptance pending

Recipe Designer selects outputs and offers only compatible **Feeds Process**
targets. Matching uses output item/SKU identity, Ingredients Assignment, and
UOM; the exact downstream `RequirementId` remains an internal bound key rather
than an operator ingredient dropdown. The full-width **Output Flow** shows
parallel stages, converging outputs, downstream output flow, and terminal
**Finished inventory** rows. Auto Order derives topological stages while
preserving unique deterministic execution ordinals. Focused RED/GREEN,
packaged evidence, and the complete regression set are GREEN. Focused source is
`8/8`; Slice 4ap `9/9`, Slice 4ao `7/7`, Slice 4an `7/7`, and Production layout
`8/8` remain GREEN. Packaged Production plus restart is `2/2`; packaged XLAM
`81/81`, Ribbon/compile `142/142`, live roles `47/47`, ordered Release 1
`30/30`, dedicated NAS `16/16`, deterministic static `19/19`, and reviewed
growth `13/13` are GREEN. Visible sample-Recipe acceptance remains pending.

### Slice 4ar output-yield clarity and multi-Process run plan: packaged GREEN, visible acceptance pending

Recipe Designer's Output Flow columns are **Stage / Produced by / Output /
Feeds Process / Output Qty / Yield % / UOM**. Output Qty and Yield % always
describe the producing Process output. The editor immediately below uses
**Required Qty / Required % / UOM** for the separate downstream connection.
Quantity-defined outputs display and retain a default 100% yield and their Qty
as Yield basis quantity; explicit yield values remain editable and persistent.

Process Designer and Ingredients Assignment list boxes receive aligned headers.
Acceptable Items displays managed item name, UOM, and item/SKU code rather than
an opaque Requirement ID/code pair. Production Run - List labels its full graph
projection **Multi-Process Run Plan**, identifies inventory choices by Process
name, and exposes the existing Process selector as a filter. A disabled **Scale
from target output Qty (coming later)** option is visible as a deliberate stub;
it has no handler that changes run state in Slice 4ar.

Automated evidence recorded 2026-08-30: focused Slice 4ar RED `0/8` then
GREEN `8/8`; prior Output Flow contract `8/8`; Production layout `8/8`;
packaged public Production actions `2/2`; packaged XLAM regression `81/81`;
Ribbon/VBA compile `142/142`; live roles `47/47`; ordered Release 1 `30/30`;
dedicated NAS `16/16`; deterministic static `19/19`; and reviewed growth
`13/13`. Visible acceptance of the sample four-Process workflow remains open.

### Slice 4as compact Process Output editor: packaged GREEN, visible acceptance pending

Process Designer's visible Output editor is one compact row ordered **ID /
Output / Design / Ver / Output Qty / Yield % / Yield basis / UOM**. The hidden
output SKU control retains state without reserving a visible gap. Design,
Version, Qty, Yield %, and Yield basis move left into that space; UOM moves up
to the same row and is a dropdown-list populated from the current warehouse
**Recipe UOM Catalog**. Existing Add/Update handlers remain the action boundary.

Automated evidence recorded 2026-08-30: focused RED `1/6` then GREEN `6/6`;
packaged public Production actions and clean restart `2/2`; packaged XLAM
`81/81`; Ribbon/VBA compile `142/142`; live roles `47/47`; ordered Release 1
`30/30`; dedicated NAS `16/16`; deterministic static `19/19`; and reviewed
growth `13/13`. Visible confirmation of the compact row/dropdown remains open.

### Slice 4at location-stock allocation and Capacity stub: automated GREEN, NAS rerun and visible acceptance pending

Ingredients Assignment widens its visible managed-item `System_Key` column
without changing the SKU alternative saved by **Add Acceptable**. Production
Run's **Acceptable Inventory For Run** projects one row per managed
SKU/UOM/Location/Condition stock bucket with summed availability. The bucket's
contributing exact keys remain hidden; **Apply** deterministically expands the
requested quantity across them, and Inventory Check/completion continue using
those exact keys. Receive Item Results adds a blank **Capacity (coming later)**
column. Capacity has no handler, validation, persistence, or inventory effect.

Automated evidence recorded 2026-08-30: focused RED `0/7` then GREEN `7/7`;
packaged Production real-handler and restart evidence `2/2`, including one
visible location-stock bucket and deterministic expansion across multiple
exact keys; packaged Receiving durability `1/1`, including the blank Capacity
stub; packaged XLAM `81/81`; Ribbon/VBA compile `142/142`; live roles `47/47`;
ordered Release 1 `30/30`; deterministic static `19/19`; and reviewed growth
`13/13`. Static metrics are 154 components, 5,223 procedures, and 1,050
scanner candidates. The current NAS rerun could not start because the
configured dedicated test root was unavailable; prior NAS `16/16` remains the
last verified result. Operator confirmation remains open.

### Slice 4au Receiving result projection: packaged GREEN, visible acceptance pending

Receive Item Results is constrained to ten operator-visible columns because
MSForms ListBox `.List` does not safely expose an eleventh column. The hidden
representative exact `System_Key` is held in a row-aligned form collection and
is still passed to Add Selected and Returns disposition actions. Capacity
remains the blank, inert sixth visible column. The public Receiving control
test must load at least one result row and prove the ten-column projection and
hidden-key map. All four Receiving column-header labels disable wrapping so
they cannot paint a second line over their ListBoxes.

Automated evidence recorded 2026-08-30: focused RED `0/7` then GREEN `7/7`;
packaged Receiving public-action evidence `1/1`; packaged XLAM `81/81`;
Ribbon/VBA compile `142/142`; live roles `47/47`; ordered Release 1 `30/30`;
deterministic static `19/19`; and reviewed growth `13/13`. Static metrics are
154 components, 5,224 procedures, and 1,050 scanner candidates.

### Connection progress, aggregate reference detail, and Events view: packaged GREEN

Slice 4w renders **Connecting to warehouse storage...** before the synchronous
Windows SMB authentication call and yields once so the connection form, mouse
pointer, and Excel status bar update before Windows begins the potentially slow
network handshake. The underlying `WNetAddConnection2` call remains synchronous;
on a remote NAS, Windows may still mark Excel busy until authentication returns.

Receiving's aggregate list remains a fixed-height, one-line MSForms ListBox.
Because that control cannot wrap or vary individual row height, selecting an
aggregate row now copies its complete concatenated PO/BOL/return reference into
a separate locked, multiline **Selected references** box. The box wraps and
scrolls, follows resize, and clears when staging is cleared.

Inventory Viewer has exactly two tabs: **Inventory** and **Events**. Runtime
construction reuses the TabStrip's two native pages; operator-visible `Tab1`
or `Tab2` placeholders are prohibited. Events is read-only.
Events is an operator-action log: rows correspond to meaningful control clicks,
not internal processing steps that happen while completing those clicks.
Receipts, Returns, Dumps, Box Made/Unboxed, Shipped, and Shipping **Remove**
activity come from the
published `tblInventoryEvents` snapshot projection. Current saved box
alternatives and currently held shipment rows supplement that projection as
`BOX_DESIGNED` and `SHIP_HELD`; these two supplements describe current activity,
not an immutable revision history. Slice 4x adds **Production Input Consumed**
and **Production Output Created** rows with Recipe/Process/run references after
the corresponding packaged event-publication contract is GREEN; design
lifecycle history remains outside this inventory-action projection. Slice 4al
adds **Inventory Adjustment** for the audited Admin retirement action; its
reference and details retain the selected SKU, reason, and exact-key event
evidence without returning the retired entity to active inventory levels.
An ordinary Shipping Add also writes an internal `SHIP_RESERVE` event. That
staging artifact is not shown as **Shipment Held**, because the operator did not
use Hold and its zero inventory delta would only report that nothing happened.
Actual current held-shipment supplements remain visible. Shipping **Remove**
also remains visible because it records the explicit release of locked inventory.
Event dates use readable `yyyy-mm-dd hh:mm:ss` text rather than Excel serial
numbers. **Refresh** replaces the open Events list with the newest published
projection; it does not process or modify authority workbooks.
The Events page exposes an editable **Event range** selector. **All** is the
first-use default and preserves the complete published view. **Day**, **Week**,
and **Month** apply rolling 1-, 7-, and 30-day windows; an operator may instead
type a whole number from 1 to 36500 for a custom rolling window. **Refresh** applies
the selected range together with the Search text. The control is hidden on the
Inventory page, so inventory-level filtering remains text-only.
Each valid applied range is stored through the existing per-Windows-user invSys
preference convention and restored when a new Viewer form opens, including in a
later Excel session. Invalid stored text falls back to **All**. This preference
never enters Config, Inventory, Snapshot, or another warehouse workbook.

Receiving, Returns, Production, and Shipping now render a form-owned pending
status and repaint before persistence starts. Excel-native Saving windows cannot
be moved into a UserForm and may still appear during required NAS saves; invSys
does not suppress required durability merely to hide Office progress UI.

### Production resizing: responsive Run List packaged GREEN; visible retest pending

`frmProduction` maximized its native window while the MultiPage and child
controls remained near their base dimensions in the upper-left corner. The
focused validator reproduced the failure at DPI-derived `Zoom=60`. The deployed
correction holds the form at `Zoom=100`; packaged validation now passes native
minimize/restore/maximize/restore and bounds/fill/overlap checks for all five
pages.

Slice 4ay adds responsive readability to **Production Run - List** without
changing any run, inventory, identity, or completion contract. The form's
readable baseline is `1110 x 800` points. **Inventory Check** reserves at least
eight rendered rows and **Selected Process Instructions** at least four. The
Recipes, **Multi-Process Run Plan**, **Acceptable Inventory For Run**,
**Inventory Check**, **Selected Process Instructions**, and **Production
Output** lists share added/removed Run-List height through Operations-local
proportional anchors. Their captions and relevant column headers follow the
same bands. **Selected Process Instructions** is Left + Right anchored, so its
left edge stays fixed while its right edge follows width changes.

Focused D13 RED was `0/5`; source GREEN is `5/5`. The packaged public
`mProduction.BtnOpenProductionForm` proof reports
`CheckEightRows=True`, `InstructionsFourRows=True`, `AllListsGrew=True`,
`InstructionLeftStable=True`, `HeadersAligned=True`, and
`GeometryHealthy=True`; it also retains the existing reusable two-batch and
Chai fork/convergence GREEN assertions. Native visual validation passed at
minimum, default, expanded, minimize, restore, and maximize states, capturing
the Production Run - List page. A user-visible drag-resize retest on the
deployed package remains required before visible acceptance is marked complete.

The 4ay follow-on corrects two visible Run List details without changing any
Production execution or inventory contract: the **Production Run - List** page
keeps a vertical scrollbar visible so constrained-height operators can reach
all of its controls, and caption anchoring explicitly excludes `hdr...` column
labels. This prevents the Run Plan **Qty** header from being assigned to the
palette-caption band during a vertical resize. The focused contract was
meaningfully RED at `5/7` and GREEN at `7/7`; the packaged public launcher now
reports `RunListVerticalScrollbar=True` and
`RunPlanQtyHeaderAligned=True`. Native default and expanded captures show the
visible scrollbar and retained **Qty** header. The deployed-package operator
drag-resize retest remains the only outstanding visible evidence for this
layout refinement.

### Production actual-output regulation: approved controls contract; implementation pending

Slice 4az adds a planned sixth `mpProduction` page, **Production Settings**.
It will present draft-only Process and Recipe scopes so an operator can select
an output and define **Regulated**, **Floor**, and **Ceiling** in that output's
UOM. Recipe scope will either inherit the selected exact Process-version/output
default or retain an immutable Recipe-version override. It is not a mutable
warehouse preference and cannot alter an active run or an already released
definition.

**Actual Output** remains the authoritative quantity created under the new
`System_Key`; **Planned** remains comparative. The visible effective minimum
will communicate `max(scaled Floor, routed commitment)`, while the configured
ceiling remains an upper bound. Thus a configured `600–610 LB` range alongside
a `610 LB` routed requirement permits only a compliant actual at or above 610;
no form control will silently turn a 605-LB exact key into 610 LB or derive an
input/output mass balance. This section records the approved contract only;
controls and handler evidence are not yet implemented.

### Shipping status/message resizing: packaged correction GREEN; visible retest pending

The 2026-08-08 screenshots show `txtStatus` initially above Search Boxes and
then translated below Search Boxes after a height resize. It was incorrectly
anchored Left + Right + Bottom. It is now anchored Left + Top + Right, so its
Top and Height remain fixed while its Width follows the form. The focused
source contract and the packaged public-launcher grow/shrink proof are GREEN.
A visible grow/shrink/maximize/restore retest remains pending.

### Inventory Viewer: Inventory and Events views packaged GREEN

Operations now contains a signed-in, capability-neutral **Inventory Viewer**
entry. It opens one reusable modeless form and reads only the current published
inventory snapshot on explicit refresh. Its Events tab reads a published event
projection plus current Shipping design/hold activity. It does not open a Receiving,
Production, or Shipping form and never writes, repairs, processes, or refreshes
an authority workbook. Packaged validation proved repeated launch reused the
same form generation, local filtering reduced the list to the matching row,
three snapshot levels loaded, Event ranges produced the expected Day, Week,
Month, custom 14-day, and All row counts, and the inspected snapshot hash did
not change.
The published projection excludes internal `SHIP_RESERVE` rows from the
operator-facing list; therefore ordinary Shipping Add does not produce a
misleading zero-quantity **Shipment Held** row. Actual held-shipment supplements
and explicit Shipping **Remove** activity remain visible.
The Operations ribbon entry now uses a visible built-in Excel table icon.

### Shipping visible identity path: `System_Key` correction packaged GREEN

`frmShipmentsTally` now exposes `System Key` in the shippable list, line editor,
shipment list, and hold list. The form passes the exact string key to the
captured-workbook staging, reservation, hold/return, and Shipments Sent paths;
the event creator emits `System_Key` without a managed `ROW` field. The focused
source contract is 12/12 GREEN, and the packaged public Shipping launcher proof
confirms that the control preserves a nonnumeric key and that the reservation
key is derived from it. Private legacy worksheet-maintenance routines are not
represented as Release 1 controls or authority paths; they remain subject to
separate reachability review before removal.

## 3. Complete form inventory

| Package/source area | Form | Current status | Control surface |
|---|---|---|---|
| Admin | `frmAddInventoryItem` | Active, runtime-generated | Add/edit inventory metadata and custom fields |
| Admin | `frmAdminDesignLifecycle` | Active, runtime-generated | Design refresh/import/release/obsolete actions |
| Admin | `frmAdminSettings` | Active, runtime-generated | Config, connection policy, carriers, UOM catalog |
| Admin | `frmCreateDeleteUser` | Active, runtime-generated | Warehouse user and capability maintenance |
| Admin | `frmCreateWarehouse` | Active, designer-backed plus runtime buttons | Greenfield warehouse creation |
| Admin | `frmReAuthGate` | Active, designer-backed | Destructive-action re-authentication |
| Admin | `frmRetireMigrateWarehouse` | Active, designer-backed plus runtime controls | Archive, retire, migrate, or delete warehouse |
| Admin | `frmSeedInventory` | Active, runtime-generated | Choose, seed, delete, or upload a demo inventory data set |
| Admin | `frmSetupTesterStation` | Active, reused designer plus runtime controls | Disposable tester-station setup and cleanup |
| Core | `frmItemSearch` | Active runtime canvas | Controls and role profiles supplied by `cDynItemSearch` |
| Core | `frmSignIn` | Active, runtime-generated | invSys user authentication |
| Core | `frmWarehouseConnection` | Active, runtime-generated | NAS/root connection and warehouse target selection |
| Operations | `frmInventoryViewer` | Active, runtime-generated | Read-only current inventory levels, search, freshness, refresh |
| Production | `frmProduction` | Active, runtime-generated; Slices 4x/4y GREEN, Slice 4z packaged GREEN with visible UAT pending | Process design, multi-table worksheet import, Recipe graph design, ingredient assignment, run-list, and experimental run-tree workflows |
| Receiving | `frmReceiving` | Active, runtime-generated | Receiving, outbound Return/Dump disposition, and non-operational Purchasing tab |
| Shipping | `frmBoxVersionSaveChoice` | Active, runtime-generated | Choose update-versus-new box alternative behavior |
| Shipping | `frmShipmentsTally` | Active, runtime-generated | Shipping, Box Designer, and Box Maker tabs |

## 4. Admin forms

### 4.1 `frmAddInventoryItem` — Add/Edit Inventory Item

This form edits inventory item metadata and queues the supported add/edit
workflow. It does not make a worksheet row the durable identity.

| Control | Type / displayed text | Purpose |
|---|---|---|
| `btnDeleteItem` | Button — **Delete Item** | Visible only in Edit mode. After an exact SKU is selected, requires confirmation and a reason, then retires all active exact `System_Key` entities for that SKU through audited Admin adjustment lines. Counted balances become zero; non-counted Utility/Service entities receive zero-delta retirement evidence. Canonical catalog/entity/event history is retained, while the retired item is removed from ordinary Edit search, managed inventory, Viewer levels, and Production pickers after Refresh. |
| `btnAddMode` | Button — **Add Item Mode** | Selects add mode and clears add fields. |
| `btnEditMode` | Button — **Edit Item** | Selects edit mode and exposes inventory search/results. |
| `btnCreateInventoryTable` | Button — **Create Inventory Table** | Creates a uniquely named bulk ADD/EDIT staging table in the captured workbook and closes the modal form so the worksheet can be edited. |
| `btnUploadInventoryTable` | Button — **Upload Selected Inventory Table** | Resolves the selected invSys inventory table in the captured workbook, preflights every pending row, then submits valid rows through existing Admin authorities and writes row outcomes. |
| `cmbEditItem` | Combo box — Inventory item | Typed text filters candidates. Choosing a dropdown row binds its exact catalog SKU and loads that item's editable fields. Visible display text alone is not identity. |
| `lstEditItemResults` | List box | Shows filtered candidate inventory items; choosing a row converges on the same SKU-backed edit loader as a combo dropdown choice. |
| `txtItemName` | Text box — Item name * | Required descriptive item name. |
| `cmbUom` | Combo box — UOM * | Selects the unit of measure. |
| `txtQty` | Combo box — Starting qty * / Set qty / Qty mode | For counted Add, accepts a required numeric Starting Qty of zero or greater; negative values are rejected. Zero creates an active managed item before stock arrives or is produced. Edit continues to accept a nonnegative target quantity. **Utility**, **Service**, and **Not counted** remain available for non-counted catalog items; Utility saves `TRACK_QTY=FALSE` and `ITEM_KIND=UTILITY`. |
| `cmbLocation` | Editable combo box — Default location | Selects a distinct current catalog location or the configured warehouse default; also accepts a new location value. |
| `cmbCategory` | Editable combo box — Category | Selects a distinct current catalog category or accepts a new category value. |
| `txtDescription` | Text box — Description | Stores the long description. |
| `txtVendorName` | Text box — Vendor(s) | Stores vendor display metadata. |
| `txtVendorCode` | Text box — Vendor code | Stores the vendor-specific code. |
| `txtExternalCode` | Text box — External code | Stores a non-invSys external code. |
| `txtImagePath` | Text box — Picture path/URL | Stores an image reference. |
| `txtEditReason` | Text box — Why the edit? * | Requires an audit reason in edit mode. |
| `txtCustomName` | Text box — Additional field | Names a proposed custom field. |
| `txtCustomValue` | Text box — Value | Supplies the proposed custom-field value. |
| `btnAddField` | Button — **Add** | Adds the name/value pair to the pending custom-field list. |
| `lstCustomFields` | List box | Displays pending custom fields. |
| `btnRemoveField` | Button — **Remove** | Removes the selected pending custom field. |
| `btnOK` | Button — **Add Item** or edit action | Validates and submits the current mode. |
| `btnCancel` | Button — **Cancel** | Closes without submitting. |
| `lblTitle` | Label | Shows the current add/edit title. |
| `lblContext` | Label | Explains that invSys generates the internal item code. |
| `lblGenerated` | Label | Shows generated identifier/result context. |
| `lblStatus` | Label | Displays validation and completion status. |
| `lblEditItem`, `lblItemName`, `lblUom`, `lblQty`, `lblLocation`, `lblCategory`, `lblDescription`, `lblVendorName`, `lblVendorCode`, `lblExternalCode`, `lblImagePath`, `lblEditReason`, `lblCustomName`, `lblCustomValue` | Field labels | Identify the controls listed above. |

### 4.2 `frmAdminDesignLifecycle` — Designs Lifecycle

The Admin ribbon exposes one **Design Lifecycle** launcher. Release and
Obsolete are actions inside this form, not separate ribbon launchers.

| Control | Type / displayed text | Purpose |
|---|---|---|
| generated label | **Design versions** | Section heading. |
| `lstDesigns` | Multi-column list box | Lists design/version/status records. |
| `btnRefresh` | Button — **Refresh** | Reloads the design list. |
| generated label | **Audit note** | Identifies the required note field. |
| `txtNote` | Multiline text box | Captures the release/obsolete audit note. |
| `btnRelease` | Button — **Release selected** | Requests release of the selected design version. |
| `btnObsolete` | Button — **Obsolete selected** | Requests obsolescence of the selected version. |
| `btnMigrate` | Button — **Import Legacy Recipes** | Invokes the explicit recipe-migration action. This is not old inventory migration. |
| `lblStatus` | Label | Displays action results or errors. |
| `btnClose` | Button — **Close** | Closes the form. |

### 4.3 `frmAdminSettings` — Warehouse Settings

| Area | Controls | Purpose |
|---|---|---|
| Heading/context | `lblTitle`, `lblConfigWorkbook`, `lblStatus` | Shows **Warehouse Settings**, the current config workbook, and action status. |
| Config grid | `lblConfigKeyHeader`, `lblConfigValueHeader`, `lblConfigTypeHeader`, `lblConfigScopeHeader`, `lblConfigRequiredHeader`, `lstConfig` | Lists config key, value, type, scope, and required status. |
| Selected config | `lblSelectedKey`, `txtConfigKey`, `lblSelectedValue`, `txtConfigValue` | Shows the selected key and permits value editing. |
| Config actions | `btnSaveConfig` (**Save Value**), `btnReloadConfig` (**Reload**) | Saves the selected value or reloads config. |
| Connection policy | `lblServerConnection`, `chkManualServerCredentials`, `lblServerConnectionScope`, `btnSaveConnectionPolicy` | Controls whether Connect Server always requests manual credentials for this Windows user. |
| Carrier catalog | `lblSection`, `lblCarrier`, `txtCarrier`, `btnAdd`, `btnRemove`, `btnReset`, `lstCarriers` | Adds, removes, resets, and lists Shipping carriers. |
| UOM catalog | `lblUomSection`, `lblUom`, `txtUom`, `btnUomAdd`, `btnUomRemove`, `btnUomReset`, `lstUoms` | Adds, removes, resets, and lists recipe UOM values. |
| Exit | `btnClose` — **Close** | Closes the settings form. |

### 4.4 `frmCreateDeleteUser` — Users & Roles

| Area | Controls | Purpose |
|---|---|---|
| Title/status | `lblTitle`, `lblStatus` | Shows **Users & Roles** and workflow status. |
| Warehouse root | `lblRoot`, `txtWarehouseRoot`, `btnRootFind`, `btnRootScan` | Selects or scans the warehouse root. |
| NAS credentials | `lblNasUser`, `txtNasUser`, `lblNasPassword`, `txtNasPassword`, `btnNasConnect`, `btnNasLogout` | Connects or disconnects this Windows session from the NAS. Password text is masked. |
| Warehouse selection | `lblWarehouse`, `cmbWarehouse`, `lblAuthPath`, `txtAuthPath`, `lblWarehouses`, `lstWarehouses` | Selects a discovered warehouse and shows its Auth workbook path. |
| User list | `lblUsers`, `lstUsers`, `btnRefreshUsers` | Lists and reloads users in the selected Auth workbook. |
| User identity | `lblUserId`, `txtUserId`, `lblDisplay`, `txtDisplayName` | Edits user ID and display name. |
| Credential | `lblPin`, `txtPin`, `btnGeneratePin`, `btnCopyPin` | Accepts or generates a PIN/password and copies the generated value when allowed. |
| Scope | `lblWhScope`, `txtWarehouseId`, `lblStationScope`, `txtStationId`, `chkRoamingStation` | Sets warehouse and station scope; Roaming removes the fixed-station requirement. |
| Capabilities | `lblRoles`, `chkAdmin`, `chkReceivePost`, `chkReceiveView`, `chkShipPost`, `chkProdPost`, `chkInboxProcess` | Assigns Admin maintenance, Receiving post/view, Shipping post, Production post, and inbox-processor capabilities. |
| Record actions | `btnClear`, `btnSave`, `btnDeactivate`, `btnDelete` | Clears the editor, creates/updates a user, deactivates the selected user, or deletes it through the guarded action. |
| Exit | `btnClose` | Closes the form. |

### 4.5 `frmCreateWarehouse` — Create Warehouse

Designer-backed controls:

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblIntro` | Intro label | Explains greenfield warehouse creation and initial publication. |
| `txtWarehouseId` | Text box — Warehouse ID | Required stable warehouse identifier. |
| `txtWarehouseName` | Text box — Warehouse Name | Required display name. |
| `txtStationId` | Text box — Station ID | Creates the initial station, defaulting to the Windows computer name. |
| `txtAdminUser` | Text box — Admin User | Creates/scopes the initial admin user. |
| `txtPathLocal` | Text box — Local Path | Target warehouse hub/runtime path. |
| `txtPathSharePoint` | Text box — SharePoint Path | Locally synced invSys SharePoint root. |
| `chkPublishInitial` | Check box — Publish initial artifacts to SharePoint | Enables initial discovery/artifact publication. |
| `lblSummary` | Wrapped status label | Shows validation, bootstrap, publication, and generated-workbook results. |
| `btnOK` | Button — **Create**, **Retry Publish**, or **Close** | Advances creation/retry/completion according to form state. |
| `btnCancel` | Button — **Cancel** or **Close** | Cancels before creation or closes after completion. |
| `lblWarehouseIdError`, `lblWarehouseNameError`, `lblStationIdError`, `lblAdminUserError`, `lblPathLocalError`, `lblPathSharePointError` | Inline error labels | Show validation errors beside their fields. |
| `lblWarehouseIdCaption`, `lblWarehouseNameCaption`, `lblStationIdCaption`, `lblAdminUserCaption`, `lblPathLocalCaption`, `lblPathSharePointCaption` | Field labels | Identify the six input fields. |

Runtime-added controls:

| Control | Displayed text | Purpose |
|---|---|---|
| `btnHubPathHelperRuntime` | **Browse...** | Browses for the NAS hub/runtime folder. |
| `btnSharePointHelperRuntime` | **Find...** | Detects or browses for the locally synced SharePoint root. |
| `btnOpenReceivingRuntime` | **Open Workbook** | Opens the generated station-local Receiving workbook after successful creation. |

### 4.6 `frmReAuthGate` — Re-Authenticate

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblPrompt` | Label | Explains that credentials must be re-entered. |
| `lblAdminUserCaption`, `lblAdminUserValue` | Labels — Admin user | Identify and show the user being re-authenticated. |
| `lblRoleCaption`, `lblRoleValue` | Labels — Role | Identify and show the required role/capability context. |
| `lblPasswordCaption`, `txtPassword` | Label and masked text box — Password | Accept the re-authentication secret. |
| `lblError` | Error label | Displays failed validation/authentication. |
| `btnOK` | Button — **OK** | Attempts re-authentication. |
| `btnCancel` | Button — **Cancel** | Cancels the guarded action. |

### 4.7 `frmRetireMigrateWarehouse` — Retire / Migrate Warehouse

| Area | Controls | Purpose |
|---|---|---|
| Title/instructions | `lblTitle`, `lblSelectionIntro` | Explains source warehouse, operation mode, and archive settings. |
| Runtime root | `lblWarehouseRootRuntime`, `txtWarehouseRootPathRuntime`, `btnWarehouseRootBrowseRuntime`, `btnWarehouseRootRefreshRuntime`, `lblWarehouseRootErrorRuntime` | Selects/scans the NAS hub or warehouse runtime root and reports errors. |
| NAS credentials | `lblNasUserRuntime`, `txtNasUserRuntime`, `lblNasPasswordRuntime`, `txtNasPasswordRuntime`, `btnNasConnectRuntime` | Connects the current Windows session to the selected NAS root. |
| Discovered warehouses | `lblFoundWarehousesRuntime`, `lstSourceWarehousesRuntime` | Lists WarehouseId and folder for discovered runtimes. |
| Source/target | `lblSourceWarehouse`, `cmbSourceWarehouse`, `lblSourceWarehouseError`, `lblTargetWarehouse`, `cmbTargetWarehouse`, `lblTargetWarehouseError` | Chooses the source warehouse and optional migration target. |
| Mode frame | `fraMode` | Contains the operation-mode options. |
| Active mode options | `optArchiveOnlyRuntime`, `optArchiveRetireRuntime`, `optArchiveMigrateRuntime`, `optArchiveRetireDeleteRuntime` | Select archive-only, retire/archive, migrate inventory, or irreversible retire/archive/delete behavior. |
| Hidden designer mode options | `optArchiveOnly`, `optArchiveMigrate`, `optArchiveRetire`, `optArchiveRetireDelete` | Legacy designer controls hidden in favor of the runtime option controls. |
| Archive destination | `lblArchiveDestPath`, `txtArchiveDestPath`, `btnArchiveDestBrowseRuntime`, `lblArchiveDestPathError` | Selects and validates the output archive path. |
| Publication/re-auth | `chkPublishTombstone`, `lblReAuthError`, `lblDeleteWarning` | Enables tombstone publication and shows re-auth/deletion warnings. |
| Confirmation panel | `fraConfirm`, `lblConfirmIntro`, `lblConfirmSummary`, `chkConfirmAction`, `lblConfirmError` | Presents the exact action and requires confirmation for destructive work. |
| Result panel | `fraResult`, `lblResultSummary` | Displays the final result. |
| Navigation | `btnBack`, `btnCancel`, `btnOK` | Goes back, cancels, advances/executes, or closes according to the active panel. |

### 4.8 `frmSeedInventory` — Demo Inventory

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblTitle` | Label — Demo inventory for which warehouse? | Explains the target choice. |
| `lblWarehouse`, `cmbWarehouse` | Label and combo box — Warehouse | Lists valid current target options and retains WarehouseId/station/root/status metadata. |
| `lblStation`, `txtStation` | Label and locked text box — Station | Shows the Windows computer name used automatically as station identity. |
| `lblUser`, `txtUser` | Label and text box — Admin user | Supplies the invSys user for the seed event. |
| `lblRoot`, `lblRootValue` | Labels — Runtime root | Shows the selected runtime root. |
| `lblDemoDataSet`, `cboDemoDataSet` | Label and combo box — Data set | Selects **R1 Workflow Kit (built-in)** or any CSV stored in the selected warehouse's managed data-set library. Uploaded definitions remain selectable on later launches. |
| `lblStatus` | Status label | Shows target readiness, validation, or inbox-repair result. |
| `btnRepairInboxes` | Button — **Repair Inboxes** | Ensures Receiving, Shipping, and Production inboxes exist for the station. |
| `btnSeedDemoInventory` | Button — **Seed Demo Inventory** | Applies the selected data set. Repeated application skips active item/location/condition groups, while missing or depleted groups receive new keys. |
| `btnDeleteDemoInventory` | Button — **Delete Demo Inventory** | After confirmation, depletes active `DEMO-` entities through exact-`System_Key` adjustment events; canonical history is retained. |
| `btnUploadDemoInventory` | Button — **Upload Data Set** | Chooses and validates a CSV, then copies it into the selected warehouse's `admin\demo-inventory-data-sets` library. Upload does not add inventory; reopen the form, select the stored definition, and click Seed. A same-named definition is rejected rather than overwritten. |
| `btnDeleteDemoDataSet` | Button — **Delete Data Set** | After confirmation, deletes only the selected uploaded CSV definition. It does not reverse or alter inventory already seeded from that definition. The built-in R1 kit is protected and cannot be deleted. |

Uploaded CSVs require `ITEM_CODE`, `ITEM`, `QTY`, `UOM`, and `LOCATION`.
`CONDITION`, `DESCRIPTION`, `CATEGORY`, and `VENDOR` are optional. Every item
code must begin `DEMO-`, quantities must be positive, required cells must be
complete, and duplicate item/location/condition groups are rejected before
anything is queued.

The two delete controls have deliberately different authority. **Delete Demo
Inventory** depletes active demo stock through audited inventory events but
retains canonical history and all data-set definitions. **Delete Data Set**
removes an uploaded CSV definition but leaves previously seeded stock and
inventory history unchanged.

Current acceptance status: the public Admin action contract is GREEN for
built-in selection, persistent uploaded-definition selection, repeated
idempotent seed, exact-key inventory depletion, uploaded-definition deletion,
R1 deletion rejection, selected CSV seed, repeated CSV seed, invalid-file
rejection, and canceled deletion. The
isolated packaged Receiving/Production/Box Maker/Shipping full chain is GREEN;
visible dedicated-NAS acceptance of the revised form remains required.

### 4.9 `frmSetupTesterStation` — Test Environment Setup

This is an Admin-only isolated test-environment provisioner, not the normal
station identity mechanism and not merely a diagnostics runner. It is retained
because it creates and cleans up disposable warehouse/operator fixtures used by
packaged regression tests. The form reuses the Create Warehouse designer blob. It hides all designer
controls first, then selectively reuses fields and adds new prompt/actions.

Visible input mapping:

| Prompt/control | Reused or runtime input | Purpose |
|---|---|---|
| `lblUserIdPrompt` | `txtAdminUser` | Tester invSys UserId. |
| `lblPinPrompt` | `txtWarehouseName` | PIN input; the reused field is masked. |
| `lblConfirmPinPrompt` | `txtConfirmPinRuntime` | Confirms the PIN; masked. |
| `lblWarehousePrompt` | `txtWarehouseId` | Tester warehouse identifier. |
| `lblStationPrompt` | `txtStationId` | Test station identifier; ordinary workflows use the Windows computer name. |
| `lblPathLocalPrompt` | `txtPathLocal` | NAS warehouse hub path. |
| `lblSharePointPrompt` | `txtPathSharePoint` | Locally synced SharePoint root. |

Other visible controls:

| Control | Displayed text | Purpose |
|---|---|---|
| `lblAdminUserError`, `lblWarehouseNameError`, `lblConfirmPinErrorRuntime`, `lblWarehouseIdError`, `lblStationIdError`, `lblPathLocalError`, `lblPathSharePointError` | Inline errors | Report validation errors for the mapped fields. |
| `lblSummary` | Status label | Shows setup/delete/open progress and final report. |
| `btnHubPathHelperRuntime` | **Find...** | Browses for the NAS hub. |
| `btnSharePointHelperRuntime` | **Find...** | Detects or browses for the SharePoint root. |
| `btnDeleteGeneratedRuntime` | **Delete Generated** | Deletes only guarded tester-generated artifacts. |
| `btnOpenReceivingRuntime` | **Open Workbook** | Appears after success and opens the generated Receiving workbook. |
| `btnOK` | **Setup** or **Close** | Runs setup, then becomes Close. |
| `btnCancel` | **Close** | Closes the form. |

Hidden but still present from the shared designer blob:
`lblIntro`, all six original `*Caption` labels, `chkPublishInitial`, and the
unused original layout positions. They are not part of the visible tester
workflow.

### 4.10 Reviewed Admin cleanup

Reachability showed that `frmAdminControls`, `frmAdminEmail`, `frmEditUser`,
`ufAdminItemSearch`, and `ufDynItemSearchTemplate` had no current public
launcher or useful deployed control surface. They were removed before the
package rebuild. The nine active Admin forms above are the Release 1 set.

## 5. Core forms

### 5.1 `frmSignIn` — invSys Sign In

The form authenticates only against the current committed warehouse target.
Changing warehouse, station, or runtime root signs out the prior session so a
credential can never be silently checked against the previously selected
warehouse. The target line must show the selected Warehouse ID and this
computer's Windows name before the user enters the warehouse-specific
credential.

The ribbon entry is explicitly labeled **invSys Sign In** while signed out and
**invSys Sign Out** while signed in. invSys Sign Out clears only the invSys
identity/capability session so another user can sign in without repeating the
server credential. A disconnected invSys Sign In action does not reopen this
form; it tells the operator to use **Server Sign In** first.

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblTitle` | Label — Sign in to invSys | Form heading. |
| `lblTarget` | Context label | Shows selected warehouse and station or that no server is connected. |
| `lblStatus` | Status label | Shows authentication guidance/errors. |
| `lblUser`, `txtUser` | Label and text box — User ID | Accepts the invSys user ID. |
| `lblSecret`, `txtSecret` | Label and masked text box — PIN/password | Accepts the invSys credential. |
| `btnOK` | Button — **Sign In** | Authenticates against the selected warehouse Auth workbook. |
| `btnCancel` | Button — **Cancel** | Cancels sign-in. |

### 5.2 `frmWarehouseConnection` — Connect Warehouse Storage

| Area | Controls | Purpose |
|---|---|---|
| Heading/status | `lblTitle`, `lblStatus` | Shows **Connect Warehouse Storage** and connection guidance/results. |
| Root | `lblRoot`, `txtRoot`, `btnScan` | Accepts and scans a NAS/server warehouse root. |
| Credentials | `lblUser`, `txtUser`, `lblPassword`, `txtPassword`, `btnConnect` | Connects the Windows session to the server; password is masked. |
| Station | `lblStation`, locked `txtStation`, `lblStationHelp` | Shows the Windows computer name. The role contract still decides whether a station inbox is required; the user does not choose a bespoke station. |
| Targets | `lblTargets`, `lstTargets` | Lists discovered warehouse runtimes and their readiness. |
| Confirmation | `btnOK`, `btnCancel` | Selects the highlighted target or cancels. |

The Operations **Send To** selector and this form share the same target-binding
contract. If the selected warehouse still contains only a legacy station such
as `S1`, selecting it may add this computer's Windows name to station config
after the warehouse Auth workbook validates. Arbitrary typed or stale station
names are still rejected. The selected target is committed only after all
validation succeeds; a failure reports the attempted target and explicitly
states that the prior current target remains unchanged.

Automated 2026-08-17 evidence is GREEN for the exact ribbon action path,
current-computer enrollment, cross-target sign-out, authentication binding, and
two clean dedicated-NAS sessions against `WHT7025AE` / `X1-PRO-AI`. The visible
operator sign-in with the user's existing warehouse credential remains the next
acceptance action.

The separate server control reads **Server Sign In** while disconnected and
**Server Sign Out** while connected. Server Sign Out also clears invSys
authentication and the selected target, disconnects the Windows SMB session,
returns both controls to their Sign In labels, and leaves capability-gated
operator buttons disabled. Selecting a warehouse through **Send To** now forces
Excel to requery the deployed Operations dropdown and server/access labels
immediately; opening Runtime Context is not required to refresh them.

### 5.3 Dynamic item-search window (`cDynItemSearch`)

The controller adds the following four controls to the clean Core-owned
`frmItemSearch` runtime canvas. Callers select a role profile rather than a
role-named form:

| Control | Type | Purpose |
|---|---|---|
| `txtSearch` | Text box | Filters the loaded role/mode-specific list. |
| `chkShippable` | Check box — **Shipping items only** or **Shippable only** | Restricts results for Shipping/component modes. |
| `lst` | Multi-column list box | Displays search results; selection/double-click/keyboard behavior routes the chosen item. |
| `txtDesc` | Read-only multiline text box | Displays description/details for the current result. |

The form caption varies by context: Receiving, Shipping, Production, Admin,
Production Ingredient, Production Palette Item, Production Recipe, or Shipping
Component Search.

### 5.4 Shared runtime canvas

`frmItemSearch` intentionally contains no designer controls or legacy event
code. `cDynItemSearch` owns all controls, search behavior, result routing, and
role-profile wording. This is the only Release 1 item-search form component.

## 6. Operations forms

### 6.1 `frmInventoryViewer` — Inventory levels and events

Ribbon entry: Operations > Overview > **Inventory Viewer**. It is visible
without a role capability restriction, while its action requires a signed-in
invSys session and selected warehouse.

| Control | Type / displayed text | Purpose |
|---|---|---|
| `tabsInventoryViewer` | Tabs — **Inventory**, **Events** | Switches between the current inventory-level projection and the read-only event/activity projection. |
| `lblTitle` | Label — **Current inventory levels** | Identifies the read-only overview. |
| `btnRefresh` | Button — **Refresh** | Reads the current published inventory snapshot or Events projection; it does not process or alter authority workbooks. |
| `lblSearch`, `txtSearch` | Label and text box — **Search** | Filters the already loaded rows locally across all visible columns. |
| `lblHeaders` | Header label | Inventory identifies Item Code, Item, UOM, Quantity, Location, and Condition. Events identifies Date, Event, Reference, Item, Qty, UOM, Location, Condition, User, and Details. |
| `lstInventory` | Six- or ten-column list box | Inventory displays active managed levels aggregated by item code, item, UOM, location, and condition, including active zero-quantity items created before stock arrives or is produced; retired items remain omitted. Events currently displays Receipts, Returns, Dumps, Box Made/Unboxed, Shipped, **Remove** (the `SHIP_RELEASE` action that returns locked inventory to warehouse availability), **Inventory Adjustment** (including audited Admin item retirement), current Box Designs, and current Held Shipments. Slice 4x adds Production Input Consumed and Production Output Created after packaged GREEN. Both views are read-only. |
| `lblStatus` | Status/freshness label | Shows row count, snapshot read time, or a no-snapshot/sign-in error. |
| `btnClose` | Button — **Close** | Closes the Viewer without affecting an operator workbook. |

The modeless form is reused on repeated launch. It is resizable: Search,
headers, list, and status expand or reposition through the shared Operations
anchor manager while remaining readable.

## 7. Receiving forms

### 7.1 `frmReceiving` — Receiving, outbound disposition, and Purchasing stub

| Area | Controls | Purpose |
|---|---|---|
| Tabs | `tabsReceiving` — **Receiving**, **Returns**, **Purchasing** | Switches between ordinary receipts, outbound inventory disposition, and the reserved Purchasing page. Receiving and Returns share one captured-workbook staging surface; Purchasing has no write action. |
| Receipt identity | `lblReceiptId`, `txtReceiptId` | Shows a generated, locked receipt ID. |
| Reference | `lblRef`, `txtRef` | Accepts PO/BOL reference text on Receiving or a Disposition Ref on Returns. |
| History filter | `lblSearch`, `txtSearch` | Filters Receiving Entries History as the user types. |
| Managed item search | `lblItemSearch`, `txtItemSearch`, `lblReceiveItemsTitle`, `lblReceiveItemsHeader`, `lstReceiveItems` | Filters and displays exactly ten visible result columns: Code, Item, UOM, Available, Location, blank **Capacity (coming later)**, Lot, Condition, Description, and Vendor. Capacity is an inert future-work stub. The representative exact `System_Key` is stored outside the ListBox in a row-aligned internal map, so identity remains available to the action handlers without exceeding the MSForms ten-column `.List` limit. The Condition column is visible on both Receiving and Returns. Rows with the same item code/UOM/location/lot/condition are aggregated and nonpositive totals are hidden. Receiving uses the hidden representative key for catalog lookup and creates a new durable inventory key. Returns uses the selection to allocate the requested quantity deterministically across the exact existing `System_Key` entities represented by that row. The header is single-line and cannot wrap over the list. |
| Quantity | `lblQty`, `txtQty` | Accepts quantity for the selected item; defaults to 1. |
| Receipt condition | `lblCondition`, `cboCondition` | Establishes physical condition at receipt-line creation. Defaults to `GOOD`; R1 choices are `GOOD`, `BAD`, `DAMAGED`, `EXPIRED`, and `REJECTED`. On Returns it is selection-backed and locked so a disposition cannot silently cross a Condition bucket. Condition is not edited in Inventory Viewer. |
| Receipt/disposition attributes | `lblReceiveLocation`, `txtReceiveLocation`, `lblLotNumber`, `txtLotNumber` | Receiving requires a location and accepts an optional lot; the operator may edit both before staging. On Returns the selected source location and lot are locked so the action preserves the exact inventory boundary. Lot is a traceability grouping, not identity or condition. |
| Disposition type | `lblDisposition`, `cboDisposition` | Visible and required only on Returns. Offers `RETURN` for goods sent out to a vendor/other party and `DUMP` for goods discarded; defaults to `RETURN`. Both reduce on-hand inventory. |
| Disposition reason | `lblReturnReason`, `txtReturnReason` | Visible and required only on Returns. Records why the selected inventory is being returned or dumped. |
| Top actions | `btnRefresh`, `btnAdd` | Reloads views or stages the selected item/quantity. The add button reads **Add Selected** on Receiving and **Add Disposition** on Returns. Both ordinary receipt staging and disposition allocation execute their complete tally/aggregate transaction with Excel events isolated, restore the caller's prior event state, and report the failing stage, error number, source, and description if Excel rejects a write. |
| Receiving/Return history | `lblInventoryTitle`, `lblInventoryHeader`, `lstInventory` | Displays a ten-column `ReceivedLog` projection: date, receipt type, reference, item, quantity, UOM, location, lot, condition, and return reason. The title is **Receiving Entries History** on Receiving and **Return Entries History** on Returns. Full user/vendor/code/`System_Key`/`EventId` evidence remains in the workbook log. |
| Staged receipt/return | `lblStagedTitle`, `lblStagedHeader`, `lstStaged` | Displays the ten-column local tally projection: reference, type, item, quantity, UOM, location, lot, vendor, condition, and return reason. The title is **Received Tally** on Receiving and **Return Tally** on Returns. This table, not the aggregate view, is the posting authority; every line retains its separate immutable `System_Key`, item code, source key, event ID, and workflow state. |
| Aggregate view | `lblAggregateTitle`, `lblAggregateHeader`, `lstAggregate` | Displays **Aggregate Received** or **Aggregate Returns**. It rebuilds from every tally row, groups by receipt type, item code, UOM, location, lot, and Condition, sums quantity, and concatenates distinct PO/BOL/return references in first-seen order. Different Conditions remain on separate rows. Return reasons are likewise concatenated for display. The aggregate is read-only and never collapses the separately keyed posting lines. |
| Aggregate reference detail | `lblAggregateReferences`, `txtAggregateReferences` — **Selected references** | Shows the complete concatenated PO/BOL/return reference for the selected aggregate row in a locked, wrapped, vertically scrollable box. This is the supported readable alternative to variable-height ListBox rows, which MSForms does not provide. It clears with the aggregate/staging surface. |
| Write actions | `btnConfirm`, `btnClear` | **Confirm Writes** queues ordinary receipts as `RECEIVE` and creates a new durable inventory `System_Key`. On Returns the button reads **Confirm Dispositions** and queues distinct `RETURN` or `DUMP` events against the staged existing keys; the Domain applies the positive action quantity as a negative exact-entity delta and rejects overdraw. **Clear** clears local staging. Multi-line confirmation batches queue and persistence work by safe workbook/artifact phase, so save cycles do not grow once per row. The real confirmation handler remains inside the shared quiet-UI boundary for queue, processor, refresh, and cleanup work, then restores the prior Excel UI/event/calculation state. A successful confirmation appends one `Persistence summary:` line to `txtStatus`; the processor remains the only snapshot writer. |
| Status/exit | `txtStatus`, `btnClose` | Shows multiline action, error, and consolidated persistence status and closes the form. Excel-native save-progress windows, if any, are separate Office UI and are not duplicated as invSys dialogs. |
| Purchasing placeholder | `lblPurchasingStub` | States that Purchasing is not operational and exposes no purchasing write action. |

### 7.2 Reviewed Receiving cleanup

`frmReceivingSavedList`, `ufReceivingItemSearch`, and
`ufDynItemSearchTemplate` were unreachable empty shells and are removed. The
intentional Purchasing stub remains inside the active `frmReceiving` because
the reviewed Release 1 workflow reserves Purchasing there.

Healthy invSys sign-in reads existing valid Config/Auth workbooks without
normalizing or saving them. Schema repair still occurs when a required table or
header is genuinely missing. This keeps the necessary repair path while
removing the repeated Saving notices caused by unchanged read-only loads.

## 8. Production forms

### 8.1 `frmProduction` — shared role form

| Control | Type | Purpose |
|---|---|---|
| `mpProduction` | MultiPage | Five pages: **Process Designer**, **Recipe Designer**, **Ingredients Assignment**, **Production Run - List**, and experimental **Production Run - Tree**. |
| `txtProductionStatus` | Locked multiline text box | Shows bound workbook, inventory/design authority, validation, action, and consolidated persistence status. |
| `btnProductionClose` | Button — **Close** | Closes the Production form. |

### 8.2 Process Designer and Recipe Designer

#### Process Designer

| Control group | Required displayed controls/actions | Purpose |
|---|---|---|
| Saved Processes | header-backed `lstProcesses`; `btnProcessRefresh`, `btnProcessNew`, `btnProcessLoad` -- **View Process**, `btnProcessReuse` -- **Edit as New Version** | Lists **ID / Version / Process / Status**. View loads an immutable saved version for reference. Edit as New Version loads its complete definition into the next generated DRAFT version and rebases retained Output Design Versions without rewriting the selected version. |
| Process identity | `txtProcessName`, locked `txtProcessId`, locked `txtProcessVersion`, `txtProcessDescription` | Edits the versioned Process header. invSys allocates the next available three-character Base-36 ID and version; the operator does not type either identity. |
| Worksheet workbench | `btnProcessWorksheetCreate` -- **Send Process to Sheet**; `btnProcessWorksheetRetrieve` -- **Retrieve Selected Process**; `btnProcessWorksheetAddAlternative` -- **Add Acceptable Item** | Sends the current new/editable Process to one of any number of uniquely named Process tables in the exact captured `Production.Operator.xlsm`. If the current definition is an existing immutable version, Send first promotes it to the next generated version. Add Acceptable Item appends one numbered managed-item/hidden-SKU pair to the selected table; retrieval accepts one table or Ctrl+click cells across several tables, imports each confirmed definition as DRAFT, and deletes only successful selected tables. |
| Requirements | header-backed `lstProcessRequirements`; locked `txtRequirementId`, `txtRequirementName`, `txtRequirementQty`, `txtRequirementPercent`, `txtRequirementYieldBasis`, `txtRequirementUom`; `btnProcessRequirementAdd`, `btnProcessRequirementUpdate`, `btnProcessRequirementRemove`, `btnProcessRequirementUp`, `btnProcessRequirementDown` | Defines typed external/upstream input requirements under **ID / Requirement / Qty / % / Batch basis / UOM**. IDs are generated Base-36 values. Worksheet input rows calculate basis and percent within each normalized-UOM group, so unlike groups may coexist without implicit conversion. |
| Outputs | header-backed `lstProcessOutputs`; locked `txtProcessOutputId`, `txtProcessOutputDesignId`, `txtProcessOutputDesignVersion`; `txtProcessOutputName`, `txtProcessOutputQty`, `txtProcessOutputPercent`, `txtProcessOutputYieldBasis`; catalog-dropdown `cmbProcessOutputUom`; hidden/internal `txtProcessOutputItemCode`; `btnProcessOutputAdd`, `btnProcessOutputUpdate`, `btnProcessOutputRemove`, `btnProcessOutputUp`, `btnProcessOutputDown` | Defines one or more output designs in one compact row under **ID / Output / Design / Ver / Output Qty / Yield % / Yield basis / UOM**. The hidden SKU reserves no visible gap. Quantity-defined outputs default to Yield %=100 and Yield basis=Output Qty; explicit percentage/basis values survive Update/save/reload. Output and Design identities are generated. UOM is selected from Settings' Recipe UOM Catalog. A managed output item is picker-selected and its SKU is retained internally; the operator does not type Item Code. Save/Release rejects a Process with no output, a noncatalog UOM, or an output without a selected managed SKU. |
| Worksheet table columns | **Record Type**, text-safe generated **ID**, **Name**, **Qty**, **Percent**, **Basis Qty**, catalog-dropdown **UOM**, generated **Design ID**, **Design Version**, **Instruction**, automatic **Requirement ID**, hidden/system-managed **Output SKU**, **Acceptable Managed Item 1** through **4** (and added pairs), with each matching managed hidden **Accepted SKU n** | Record Type is dropdown-backed. INPUT, OUTPUT, and INSTRUCTION IDs use one table-wide Base-36 namespace and remain unique regardless of entry order; INPUT Percent/Basis Qty are formula-managed per normalized-UOM group, INPUT Requirement ID and OUTPUT Design identity are system managed, and mixed-UOM assembly rows are valid when every group totals 100.0%. Core item search opens only from a valid **Acceptable Managed Item n** cell: INPUT fills that numbered alternative pair; OUTPUT pair 1 fills the visible managed item and hidden Output SKU while retaining its descriptive Name. OUTPUT Name never opens search. Neither path stores a physical `System_Key`; historical ALTERNATIVE rows remain import-compatible. |
| Instructions | header-backed `lstProcessInstructions`, `txtProcessInstruction`; `btnProcessInstructionAdd`, `btnProcessInstructionUpdate`, `btnProcessInstructionRemove`, `btnProcessInstructionUp`, `btnProcessInstructionDown` | Lists **Step / Instruction** and defines reusable operator instructions independently from input/output rows. |
| Lifecycle | `btnProcessValidate`, `btnProcessSave`, `btnProcessRelease`, `btnProcessObsolete`, `btnProcessClear` | Validates and queues immutable Designs Domain Process lifecycle events. |

#### Recipe Designer

| Control group | Required displayed controls/actions | Purpose |
|---|---|---|
| Saved Recipes | `lstRecipes`; `btnRecipeRefresh`, `btnRecipeNew`, `btnRecipeLoad` | Lists Recipe ID/version/status records and starts or loads a draft. |
| Recipe identity | `txtReusableRecipeName`, locked `txtReusableRecipeId`, editable `txtReusableRecipeVersion`, `txtReusableRecipeDescription` | Edits the versioned Recipe header. invSys allocates the next collision-checked three-character Base-36 Recipe ID and proposes version `1` for a blank draft on form load, New Recipe, Clear, and as a Save Draft/Release fallback. The operator may replace Version with another positive whole number; Recipe ID cannot be typed, and an existing immutable ID/version is never overwritten. |
| Process library/nodes | header-backed `lstReleasedProcesses`, header-backed `lstRecipeNodes`; `btnRecipeAddProcess`, `btnRecipeRemoveProcess` | Reuses exact released Process versions and assigns a hidden node identity within this Recipe. Operator projections show Process name, version/status/order rather than generated Process codes. The two lists sit side by side above Output Flow. |
| Output routing | name-visible/ID-bound `cmbConnectionFromNode`, `cmbConnectionOutput`, compatibility-filtered `cmbConnectionToNode` displayed as **Feeds Process**, hidden/internal `cmbConnectionRequirement`, `txtConnectionQty`, `txtConnectionPercent`, catalog-backed `cmbConnectionUom`, internal `lstRecipeConnections`, visible header-backed `lstRecipeConnectionDisplay`; `btnRecipeConnect`, `btnRecipeUpdateConnection`, `btnRecipeDisconnect` | Routes selected outputs to later Processes. Feeds Process contains only nodes with one unresolved requirement whose UOM and Ingredients Assignment item/SKU match the selected output; invSys binds the matching Requirement ID internally. An arbitrary downstream ingredient list is never shown and a same-node self-reference is invalid. |
| Output Flow | `lstRecipeConnectionDisplay` with **Stage / Produced by / Output / Feeds Process / Output Qty / Yield % / UOM** | Full-width projection of connected and terminal outputs using the producing Process output definition. Independent outputs can share a stage and converge on one later Process; unconnected outputs display **Finished inventory**. The separate editor labels connection values **Required Qty / Required % / UOM**. Hidden node/Output/Requirement IDs remain the stored graph identity. |
| Execution order | `lstRecipeNodes`; `btnRecipeMoveUp`, `btnRecipeMoveDown`, `btnRecipeAutoOrder` | Auto Order derives deterministic topological execution order and visible stages. Independent Processes may share a visible stage while persisted execution ordinals remain unique. Validation rejects an order inconsistent with the directed graph. |
| Validation/lifecycle | `lstRecipeValidation`; `btnRecipeValidate`, `btnRecipeSave`, `btnRecipeRelease`, `btnRecipeObsolete`, `btnRecipeClear` | Reports unresolved inputs, compatibility, quantity/yield, missing definition, and circular-dependency failures before lifecycle events are queued. |

### 8.2.1 Historical pre-Slice 4x Recipe Builder surface

These controls document retired historical behavior. `BuildRecipeBuilderPage`
is retained as unreachable legacy source for reviewed cleanup only; the runtime
five-page form does not construct it, and it is not a Designs-enabled authority
or fallback.

| Control group | Controls | Purpose |
|---|---|---|
| Saved recipes | `lstBuilderRecipes` | Lists recipe ID/name/description choices. |
| Recipe fields | `txtRecipeName`, `txtRecipeId`, `txtRecipeRowBudget`, `txtRecipeDescription` | Edit recipe identity, row budget, and description. |
| Recipe actions | `btnBuilderRefresh`, `btnBuilderNew`, `btnBuilderLoad`, `btnBuilderSave`, `btnBuilderProcess`, `btnBuilderFormulas`, `btnBuilderClear`, `btnBuilderRelease` | Refresh, start, load, save, add process tables, save formulas, clear, or release a recipe. |
| Line fields | `txtLineProcess`, `cmbLineIo`, `txtLineIngredient`, `txtLinePercent`, `cmbLineUom`, `txtLineAmount` | Edit process, input/output classification, ingredient/output/instruction, percent, UOM, and amount. |
| Line actions | `btnLineUomAdd`, `btnLineAdd`, `btnLineUpdate`, `btnLineRemove`, `btnLineMoveUp`, `btnLineMoveDown` | Add UOM, add/update/remove a recipe line, or reorder it. |
| Builder grid | `lstBuilderLines` | Lists Process, I/O, ingredient/output/instruction, %, UOM, amount, and ingredient ID. |
| Labels | Saved Recipes, Recipe Name, Recipe ID, Row Budget, Description, Process, In/Out, Ingredient / Output / Instruction, Percent, UOM, Amount, Recipe Builder Lines | Identify the controls and generated list headers. |

Historical pre-Slice 4x generated list-header control families are `hdrBuilderLines1` through
`hdrBuilderLines8`, `hdrLoaderRecipes1` through `hdrLoaderRecipes3`,
`hdrLoaderLines1` through `hdrLoaderLines8`, `hdrRunPalette1` through
`hdrRunPalette10`, `hdrManagerCheck1` through `hdrManagerCheck6`, and
`hdrManagerOutput1` through `hdrManagerOutput8`. Blank or hidden data columns
still received a header control where the old runtime builder created one. The
current Process/Recipe designer lists use the exact IDs recorded in section
8.2; their minimum/default/expanded and maximize/restore geometry is GREEN.

### 8.3 Ingredients Assignment page

| Control group | Controls | Purpose |
|---|---|---|
| Processes | header-backed `lstAssignRecipes`, `btnAssignRecipe`, `btnAssignRefresh` | Lists **Version / Process** and selects the exact Process version whose requirements are being assigned. |
| Requirements | header-backed `lstAssignIngredients`, `btnAssignIngredient` | Lists **Requirement / UOM / Process / Type / Qty / %** and selects one declared requirement; connected Recipe requirements do not allocate inventory alternatives during a run. |
| Assignment actions | `btnAssignSave`, `btnAssignClear` | Saves acceptable managed item/SKU alternatives as part of a new Process draft version, or clears the editor. |
| Inventory search | `txtInventorySearch`, header-backed `lstAssignInventory` | Filters and lists **System_Key / Managed Item / UOM / Inv / Location / Description** candidates with a readable-width immutable key. |
| Allowed choices | header-backed `lstAssignAllowed`, `btnAssignAdd`, `btnAssignRemove` | Lists **Managed Item / UOM / Item Code** alternatives and adds/removes rows without exposing Requirement ID or allocating a physical entity. |
| Labels | Approved wording: Processes, Ingredient Requirements, Search Inventory, Managed Items, Acceptable Items | Identifies the page sections. |

### 8.4 Production Run - List page

| Control group | Controls | Purpose |
|---|---|---|
| Recipe loader | `lstLoaderRecipes`, header-backed `lstLoaderLines` labelled **Multi-Process Run Plan**, `btnLoaderRefresh`, `btnLoaderLoad`, `btnLoaderClear` | Selects an exact released Recipe version and shows the complete validated Process graph/execution order with Process, line type, requirement/output name, scaled Qty, %, UOM, and textual Status. Status keeps **! INSUFFICIENT**, **NEEDS ALLOCATION**, **WAITING UPSTREAM**, **READY**, and **COMPLETE** visible for the whole Recipe while one Process is selected. |
| Batch scaling | `txtBatchScalePercent`, `btnApplyBatchScale` | Applies a List-run batch scale from `0.001%` through `1000%`; `100%` preserves the released recipe quantities. |
| Target-output scaling stub | disabled `chkRunTargetOutputScale`, disabled `cmbRunTargetOutput`, disabled `txtRunTargetOutputQty` | Displays **Scale from target output Qty (coming later)** without calculating or changing Recipe/run state in Slice 4ar. |
| Run inputs | `cmbRunProcess`, `cmbRunLocation`, `txtPaletteSplit`, `txtPaletteQty`, `btnRunApplyPalette` | Selects the one Process to allocate, Check In, and complete at a time, plus its run location. Applies either percent-of-requirement or explicit quantity to the selected external inventory requirement. Connected intermediate requirements become runnable only after their upstream Process output exists under its exact run key. Unresolved inputs belonging only to another Process do not block the selected Process. |
| Palette | `lstRunPalette` | Lists owning **Process**, ingredient requirement, acceptable managed stock, requirement %, allocated quantity, UOM, summed available inventory, and location while retaining hidden node/requirement and representative entity identities. One row represents one SKU/UOM/Location/Condition bucket, not one receipt. Apply expands it into exact-key allocations. At least eight ordinary result rows are visible at the approved default and minimum form sizes. |
| Run instructions | header-backed `lstRunInstructions` | Shows **Step / Instruction** for the selected Process throughout allocation, Check In, and completion. Instructions come from the pinned released Process version and are read-only in Production Run. |
| Inventory check | `lstManagerCheck` | Lists external allocations plus read-only routed inputs. Every routed row identifies the downstream requirement, source Process/output, readable exact produced `System_Key`, committed quantity, UOM, and remaining balance; it is never selectable stock and survives Process selection, Refresh, and normal navigation in the active Recipe version/RunId/batch. At minimum geometry the compact audit/output bands remain scrollable, so rows are not discarded to fit the form. External rows continue to show exact allocation identity and current available inventory; insufficiency or staleness blocks Check In/completion. |
| Outputs | `lstManagerOutput`, `txtOutputReal` -- **Actual Output** | For a reusable run, retains one row per completed batch/Process output plus the selectable active-batch row. Completed rows show **Last Actual**, Batch, **Used Goods**, cumulative **Process Total**, recall code, and their readable distinct new **System_Key**. Used Goods is a deterministic per-Process/batch grouped UOM summary such as `5 LB; 12 EA`, never a numeric sum of unlike units. Selecting an active output row loads its staged actual; editing Actual Output retains a positive per-output quantity. Complete Run requires every actual, creates managed inventory at that actual quantity, and rejects an actual smaller than routed downstream commitments. Legacy completion continues to use the same visible quantity field through its preserved path. |
| Run actions | `btnManagerCheckIn`, `btnManagerApplyOutput`, `btnManagerRefresh`, `btnManagerNext`, `btnManagerPrint` | Checks exact inputs into the selected Process, completes that Process when its dependencies are READY, refreshes, advances after every Recipe Process completes, or prints recall data. **Complete Run** keeps event/processor work inside the shared quiet-UI boundary, creates every selected-Process output under its own new key, consumes routed intermediates by exact key, and appends correlated persistence feedback. Independent READY Processes may complete in either order; downstream Processes wait for their upstream output keys. |
| Labels | Recipes, Multi-Process Run Plan, Process filter, Run Location, % of Requirement, Qty, Acceptable Inventory For Run, Inventory Check, Production Output, Actual Output | Identifies the page sections and generated headers. Exact identity remains visible in Inventory Check and Production Output but is hidden behind the Run stock-bucket projection. Output quantity columns display **Last Actual**, **Batch**, **Used Goods**, and **Process Total**. Run palette Inv and Inventory Check show **Utility** rather than a numeric balance for catalog Utility items. Normalized `EA` is whole-unit only throughout visible role flows; a fractional value is rejected rather than rounded. |

Slice 4av automated acceptance recorded 2026-08-30: the focused public-action
package completed two multi-Process batches through the operator Check In and
Complete Run handlers with `SelectedProcessOnly=True`,
`WholeRecipeStatus=True`, `RunInstructionsVisible=True`, and
`EightPaletteRows=True`. It also preserved exact-key inputs, distinct output
keys, routed-intermediate consumption, co-product balance, actual-output
inventory authority, batch history, Utility display, and stock-bucket
expansion. Visible operator acceptance of the Chai Recipe remains open.

Slice 4ax automated acceptance recorded 2026-08-31: a separate packaged public
handler run created the named four-Process **Classic Chai Fork Convergence**
fixture and completed Tea Brewing + Spice Blending -> Chai Convergence -> Final
Bottling in batch 1. The same Inventory Check list showed both read-only
routed inputs at convergence and the exact concentrate key at bottling; the
test then proved exact-key consumption, a new final bottled `EA` key, and one
unchanged Recipe/version/RunId/batch identity across Refresh and Process
selection. It is GREEN with `ChaiFourProcessesCompleted=True`,
`ChaiFinalBottlingCompleted=True`, and `ChaiRunNotRestarted=True`. This does
not replace the dedicated NAS operator-visible Chai UAT, which remains open.
The current approved five-package set was revalidated on the dedicated NAS
test leaf in two clean sessions at 16/16, with zero canonical-file changes from
launcher use; that is deployment/launcher evidence, not a visible Chai run.
That UAT must visibly preserve the displayed Recipe/version/Run/Batch across
Refresh and Process selection, show **WAITING UPSTREAM** before upstream
completion and **READY** only afterward, keep routed inputs read-only and out
of **Acceptable Inventory For Run**, and retain the exact input/output keys,
co-product balance, UOM-grouped **Used Goods**, and correlated Viewer Events
through Final Bottling.

Slice 4aw automated acceptance recorded 2026-08-30: **View Process** distinguishes
an immutable loaded definition from **Edit as New Version**. The latter keeps
the full definition but generates the next Process version and rebases each
retained Output Design Version. **Send Process to Sheet** exports the current
new/editable definition to the captured workbook; sending a viewed existing
version first promotes it to the next generated version. Retrieval saves that
edited table through the same public DRAFT authority and never rewrites the
saved/released source version. Focused source began RED at `0/6` and is GREEN
at `7/7`. The packaged public action is GREEN at `1/1` with
`ReleasedProcessEditable=True`, `ExistingProcessExported=True`,
`ExportRoundTrip=True`, `OutputDesignVersionRebased=True`, and
`OutputYieldRebased=True`. Packaged XLAM validation is `81/81`, packaged
Ribbon/VBA compile validation is `142/142`, Production layout validation passes
at minimum/default/expanded/native sizes, Slice 0 tooling is `62/62`, the full
Release 1 chain is `30/30`, and live role workflows are `47/47`. Visible
acceptance of the corrected bottling Process remains open.

### 8.5 Production Run - Tree page

This page remains an experimental alternative and is not the Release 1
workflow focus. Slice 4x intentionally makes no Tree behavior contract change;
Production Run - List is the Release 1 path to validate.

| Control group | Controls | Purpose |
|---|---|---|
| Tree choices | `lstRunTree` | Displays hierarchical recipe ingredient/inventory choices. |
| Tree inputs | `txtTreePaletteSplit`, `txtTreePaletteQty`, `cmbTreeRunProcess`, `cmbTreeRunLocation`, `btnRunTreeApplyPalette` | Applies percent/quantity/process/location to the selected tree choice. |
| Tree navigation | `btnRunTreeExpandAll`, `btnRunTreeCollapseAll` | Expands or collapses all tree groups. |
| Generated headers | Ingredient, System Key, Inventory Item, % Req, Qty, UOM, Inv, Location | Identify tree columns. |

The unconstructed `mBtnManagerPrepare`, `mBtnManagerUsed`,
`mBtnManagerMade`, and `mBtnManagerTotal` event variables/handlers were removed;
they were never visible controls. Current packaged geometry validation is GREEN
for all five pages at minimum/default/expanded size and through native
minimize/restore/maximize/restore. Visible dedicated-NAS layout acceptance
remains open.

### 8.6 Reviewed Production cleanup

`frmCreateRecipeTable`, `frmCreateSubstitutionList`, `frmIngredientPalette`,
`frmSubstitution`, `ufProductionItemSearch`, and
`ufDynItemSearchTemplate` were empty or unreachable shells and are removed.

## 9. Shipping forms

### 9.1 `frmShipmentsTally` — Shipping/Boxing role form and tabs

| Control | Type | Purpose |
|---|---|---|
| `tabsShippingRole` | TabStrip — **Shipping**, **Box Designer**, **Box Maker** | Selects the active Shipping sub-workflow. |
| `txtStatus` | Locked multiline text box | Displays validation, queue, history, and action status; remains fixed above Search Boxes during resize. |
| `btnClose` | Button — **Close** | Closes the form. |

### 9.2 Shipping tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/history | `lblTitle`, `btnHistory`, `btnHistorySheet`, `btnRefresh` | Shows Shipments, opens history, exports history to a sheet, and refreshes data. |
| Search/filter | `lblPicker`, `txtPicker`, `chkUseExisting`, `lblSyncState` | Filters boxes, optionally uses existing shippable inventory, and shows pending/complete sync state. |
| Shippables | `lstShippables` plus `hdrShipBox`, `hdrShipVersion`, `hdrShipInv`, `hdrShipProjected`, `hdrShipLocked`, `hdrShipUom`, `hdrShipLoc`, `hdrShipSystemKey` | Lists Box, Alternative, NAS Inv, Projected Inv, Locked, UOM, Location, and immutable System Key. |
| Line editor | `txtRef`, `txtBox`, `txtVersion`, `txtQty`, `txtUom`, `txtLocation`, `txtSystemKey`, `txtCarrier`, hidden `txtDescription` | Edits reference, quantity, and carrier while preserving the selection-backed box/alternative/UOM/location/System_Key identity. |
| Line labels | `lblRef`, `lblBox`, `lblVersion`, `lblQty`, `lblUom`, `lblLocation`, `lblSystemKey`, `lblCarrier` | Identify the editor fields. |
| Line actions | `btnAdd`, `btnUpdate`, `btnRemove` | Adds, updates, or removes a shipment row. Exact-key reserve/release writes remain durable; the form reports the completed inbox and reservation-ledger boundaries once in `txtStatus`. |
| Shipments list | `lblShipments`, `lstShipments` | Displays active shipment staging rows. |
| Shipment headers | generated `hdrRef*`, `hdrLineBox*`, `hdrLineQty*`, `hdrLineUom*`, `hdrLineArea*`, `hdrLineLocked*`, `hdrLineSystemKey*`, `hdrLineDesc*`, `hdrLineCarrier*` | Identify Ref, Box, Qty, UOM, Area, Locked, System Key, Alternative, and Carrier. One header set is built for each shipment/hold list. |
| Shipment actions | `btnStage`, `btnSend`, `btnHold` | Moves rows to shipment staging, sends completed shipments, or places selected rows on hold. |
| Hold list | `lblHold`, `lstHold`, `btnReturn` | Displays Not Shipped rows and returns selected rows to active staging. |

**D14 correction:** the visible form and its reachable Release 1 backing path
use exact string `System_Key`; no numeric value was relabeled as a key.
The **Add** reserve step resolves and updates that exact inventory entity on a
current schema with no managed `ROW` column.

After **Shipments Sent** processes the queued event, the form refreshes its
captured operator workbook from the processor-generated canonical snapshot,
reloads the shippables list, and then derives NAS Inv, Projected Inv, and Locked.
A completed shipment must therefore be visible immediately without depending on
an old local array or a later timer-driven refresh.

Successful Add/Update/Remove, To Shipments, and Shipments Sent reports include
one `Persistence summary:` line in `txtStatus`; the Shipments Sent modal summary
contains the same line. Multiple selected rows share one reservation-ledger
open/save cycle. The summary does not replace or weaken the canonical workbook
saves, and it does not promise suppression of Excel-native save-progress UI.

The uncalled readiness list and header builder were removed after reachability
review; they were never constructed by the active layout.

### 9.3 Box Designer tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/actions | `lblBoxBuilderPage`, `btnBoxBuilderNewPage`, `btnBoxBuilderRefreshPage` | Starts a box definition and reloads box designs. |
| Designs | `lstBoxBuilderDesignsPage` | Full-width list of design identity/name/alternative/status metadata. |
| Component inventory | `lblBoxBuilderInventory`, `lstBoxBuilderInventoryPage`, `lblBoxBuilderComponentQty`, `txtBoxBuilderComponentQty`, `btnBoxBuilderAddComponentPage`, `btnBoxBuilderRemoveComponentPage` | Selects positive-balance inventory entities and quantities for the box definition. Each row retains its hidden immutable `System_Key`; duplicate projections of the same key and nonpositive balances are omitted, while repeated item names with different keys remain distinct physical entities. |
| Component search | `txtBoxBuilderSearch` | Filters loaded component choices locally by code, item, description, UOM, or location. |
| Definition fields | `lblBoxBuilderName`, `txtBoxBuilderName`, `lblBoxBuilderVersion`, `cboBoxBuilderVersion`, `lblBoxBuilderStatus`, `cboBoxBuilderStatus`, `lblBoxBuilderUom`, `txtBoxBuilderUom`, `lblBoxBuilderLocation`, `txtBoxBuilderLocation`, `lblBoxBuilderDescription`, `txtBoxBuilderDescription` | Edits box name, alternative, Active/Archived status, UOM, location, and description. |
| Selected components | `lblBoxBuilderComponents`, `lstBoxBuilderComponentsPage` | Full-width list of the component set for the selected alternative. The hidden component identity is preserved as `System_Key` and persisted as `ComponentSystemKey` in the Shipping BOM. |
| Save/alternative actions | `btnBoxBuilderSavePage`, `btnBoxBuilderUpdateVersionPage`, `btnBoxBuilderNewVersionPage`, `btnBoxBuilderDeleteVersionPage`, `btnBoxBuilderArchivePage`, `btnBoxBuilderDeletePage` | Saves the box, updates or creates an alternative, deletes an alternative, archives the box, or deletes it through guarded paths. |

### 9.4 Box Maker tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/refresh | `lblBoxMakerPage`, `btnBoxMakerRefreshPage` | Identifies Box Maker and reloads released designs/inventory state. |
| Designs/alternative | `lstBoxMakerDesignsPage`, `lblBoxMakerVersion`, `cboBoxMakerVersion` | Selects a released box design and alternative. |
| Quantity/actions | `lblBoxMakerQty`, `txtBoxMakerQty`, `btnBoxMakerMakePage`, `btnBoxMakerUnmakePage` | Accepts quantity and makes boxes or unboxes them through the event path. |
| Components | `lblBoxMakerComponents`, `lstBoxMakerComponentsPage` | Displays selected-alternative component requirements and availability resolved by the exact preserved `ComponentSystemKey`. Make/Unmake event payloads use those string keys and the package `System_Key`, never a numeric row surrogate. |

All Designer/Maker list boxes now use an explicit vertical, full-width layout
and participate in native form resizing without overlap. Each list has
a monospaced header strip whose position and width are recalculated from the
list geometry after anchor application. A managed item without a Shipping BOM
alternative displays `NA`; only a Shipping BOM/package alternative displays
the compatible stored labels `v1`, `v2`, and so on.

### 9.5 `frmBoxVersionSaveChoice`

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblTitle` | Label — Save Box Alternative | Heading. |
| `lblBody` | Wrapped label | Explains the choice for the current box alternative. |
| `btnUpdate` | Button — **Update Alternative** | Selects in-place update behavior where allowed. |
| `btnNewVersion` | Button — **New Alternative** | Creates another box design choice; this is an alternative, not a software release sequence. |
| `btnCancel` | Button — **Cancel** | Cancels without saving. |

### 9.6 Reviewed Shipping cleanup

`frmShippingCreateList`, `frmShippingSavedList`, `ufShippingItemSearch`, and
`ufDynItemSearchTemplate` were empty or unreachable shells and are removed.

## 10. Review checklist decisions

| Item | 2026-08-08 decision / state |
|---|---|
| 1. Necessary for Release 1? | Operations/Admin and the active role-form tabs are sufficient. Unreachable shells and unconstructed controls were removed. Inventory Viewer was added as the one missing overview surface; Purchasing remains the intentional Receiving stub. |
| 2. Operator wording? | Accepted as good enough for Release 1. |
| 3. Visibility? | Controls may remain visible to all signed-in users for Release 1. Existing capability checks on mutating actions remain defense-in-depth; Inventory Viewer has no role-capability restriction. |
| 4. Mutation/queue/admin distinction? | **Release 1 decision:** Search, filters, history, Viewer Refresh, and role Refresh are read-only projections. Add/Update/Remove/Hold and disposition selection change only captured-workbook staging until an explicit post. Process/Recipe edits are local drafts until Save/Release/Obsolete queues a Designs Domain lifecycle event. Confirm Writes, Confirm Dispositions, Production completion, Make/Unmake, and Shipments Sent queue inventory events; the processor is the only inventory authority writer. `RETURN` and `DUMP` are separate audited events but share the Receiving `RECEIVE_POST` capability. Admin controls perform named administrative actions and retain capability/re-auth/audit gates. Status text should state `staged`, `queued`, `applied`, or `refreshed` rather than imply that a local edit immediately changed inventory. |
| 5. Durable identity uses `System_Key`? | **Accepted rule:** every physical inventory entity, inventory event/log row, reservation, shipment, Production allocation/output, and Shipping BOM package/component reference uses immutable `System_Key`. Every executed Process output receives a new key; an intermediate connection consumes that same exact key and never uses Process output definition identity as physical inventory identity. Aggregate quantity views may group by SKU/location and need not impersonate one key. A Return/Dump aggregate selection must be expanded into exact existing keys before posting; it never creates a replacement key or borrows quantity across item/location/lot/Condition boundaries. Process, Recipe, and other Designs definitions use their specialized version identities; an inventory entity produced from any definition still receives its own `System_Key`. `EventId`, shipment IDs, and run IDs identify their own records. |
| 6. Readable resize proportions? | Required. Production, Shipping `txtStatus`, Viewer, and the explicit full-width Box Designer/Box Maker layouts are source- and packaged-geometry GREEN. Visible maximize/restore/grow/shrink confirmation remains part of the dedicated UAT checkpoint. |
| 7. Empty shells? | Retired after source reachability review and focused regression contracts. Only the visible Purchasing stub remains intentionally. |

## 11. Required follow-up to close the control review

- Sign into Admin at `WHT7025AE` / `X1-PRO-AI` with the existing invSys user
  and confirm the Admin ribbon controls enable without a capability error.
- Repeat visible Production maximize/restore and Shipping/Boxing grow/shrink checks.
- Repeat the ordinary Receiving **Add Selected** action, then confirm one
  receipt batch and one disposition batch. Each successful confirm must show
  one `Persistence summary:` line in Receiving status; record any remaining
  Excel-native Saving notices separately.
- Complete one Production Run - List batch and confirm one `Persistence
  summary:` line appears in Production status; record any remaining
  Excel-native Saving notices separately.
- Preserve the accepted uploaded-data-set visibility in Inventory Viewer. Then
  confirm generated locked Base-36 IDs, the clarified batch/yield basis wording,
  and one complete
  **Edit Process on Sheet** / **Retrieve Process from Sheet** / repeat-edit
  round trip with formula percentages totaling 100.0%.
- After Slice 4x packaged GREEN, create/release/reuse a named multi-output
  Process, connect it in a released Recipe, confirm unresolved/circular graph
  rejection, assign acceptable alternatives, and complete two saved-workbook
  List batches with distinct output keys, exact routed-intermediate consumption,
  co-product balance, and Production Viewer Events.
- In Box Designer, confirm Component inventory omits zero-balance rows and does
  not repeat the same exact entity. Shipping Add, Update Row, Remove, Return,
  To Shipments, and Shipments Sent are visibly action-reachable, and the
  post-send NAS/Projected/Locked refresh is visibly accepted. Retest the new
  `Persistence summary:` line and record any remaining Excel-native Saving
  notices separately from the invSys status output.
- Verify Server Sign In paints its connecting status before the remote NAS wait,
  select a four-reference Aggregate Received row and read all references in
  **Selected references**, then open Inventory Viewer > **Events** and confirm
  receipt/disposition/boxing/shipping activity is visible without changing
  canonical inventory.
- Run the complete 24-row seed-to-ship workflow on the dedicated test warehouse,
  including Receiving Location/Lot/Condition, one `RETURN`, one `DUMP`, their
  resulting Viewer quantities, aggregate totals, and Production List scaling.
- Review and adjust checklist item 4 only if an operator test exposes an
  ambiguous status or mutation boundary.
- Review private legacy Shipping worksheet-maintenance routines for reachability
  before removal; they are outside the packaged Release 1 form/action authority
  path and must not be reintroduced as compatibility behavior.
