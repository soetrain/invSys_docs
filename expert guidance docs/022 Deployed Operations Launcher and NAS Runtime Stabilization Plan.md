# Deployed Operations Launcher and NAS Runtime Stabilization Plan

**Status:** Current corrective implementation plan
**Created:** 2026-07-28
**Applies to:** invSys Architecture v4.11, especially D2, D3, D12-D14, the
operator deployment model, and Phase 6 saved-workbook acceptance
**Scope:** Packaged Operations ribbon launchers, station-local role workbook
provisioning, deployed-package diagnostics, and dedicated NAS test-runtime
acceptance, including the D15 reusable-Process Production priority change

## 1. Purpose and authority

This plan follows completed plan 020. It addresses defects found in an actual
server-connected, signed-in operator session after the Release 1 isolated
full-chain gate passed.

Authority remains:

1. `invSys-Design-v4.11.md` is normative.
2. D12 keeps Receiving, Production, and Shipping in
   `invSys.Operations.xlam`.
3. D13 requires meaningful behavioral RED before implementation.
4. This document sequences assessment before implementation; the final code
   scope is evidence-driven.
5. User-observed deployed behavior is acceptance evidence, not permission to
   mutate operational NAS workbooks during diagnosis.

The Receiving decision recorded here is binding: when the Receiving control is
used without an eligible Receiving operator workbook, invSys creates or opens
the user's station-local Receiving operator workbook and then opens the form.
This is consistent with the normative station-bootstrap requirement that each
station own a local role operator workbook. It does not make the operator
workbook canonical and does not place it in the warehouse-authoritative NAS
runtime.

The user-directed Production redesign recorded in Slice 4x is a deliberate
priority and contract change. It supersedes the former single Recipe Builder
surface and singular-output run model only after the Architecture v4.11 D15
contract is updated. It does not reopen completed launcher, Seed, Receiving,
Shipping, Boxing, or Viewer behavior. Those GREEN contracts remain mandatory
regressions while Production moves to reusable Processes and Recipe graphs.

## 2. Observed deployed defects

Observed 2026-07-28 after the operator connected to the server and signed in to
invSys:

| Operations control | Observed result | Required result |
|---|---|---|
| Receiving | `Open a Receiving operator workbook before using the Receiving form.` | Create or open the station-local Receiving operator workbook, bind the modeless form to it, and leave canonical NAS workbooks untouched. |
| Production | `Production form failed: Type mismatch` | Create, open, or reuse the station-local Production operator workbook and open one captured modeless form without a VBA type mismatch. |
| Shipping | `SHIPMENTS failed: Type mismatch` | Create, open, or reuse the station-local Shipping operator workbook and open one captured modeless form without a VBA type mismatch. |

The exact `Err.Number`, `Err.Source`, failing initialization stage, loaded XLAM
paths/hashes, active workbook classification, selected warehouse target, and
resolved runtime root were not captured by the existing messages. Root cause
for the two type mismatches is therefore unresolved.

## 3. Verified coverage gap

Plan 020 Slice 14 used:

- [x] the five built XLAMs from repository `deploy/current`;
- [x] packaged Admin, Receiving, Production, Boxing, and Shipping action
  boundaries; and
- [x] a fresh generated runtime under
  `%TEMP%\invsys-release1-chain-<guid>`.

It did not use:

- [ ] the operator's actually loaded/installed XLAM paths;
- [ ] the selected UNC/NAS warehouse runtime;
- [ ] the real station-local operator-workbook discovery/provisioning path; or
- [ ] the three visible Operations ribbon controls in the connected operator
  session.

The prior 30/30 full-chain result remains valid for its isolated scope. It is
not evidence that the deployed NAS-connected launcher path passes.

## 4. Assessment before implementation

Do not choose a patch or refactor strategy from the three visible messages
alone. Before editing runtime code:

- [x] complete the read-only deployed-runtime entry gate;
- [x] reproduce all three failures through the packaged callbacks;
- [x] capture the exact failing stage, error number, sanitized source, and
  description;
- [x] compare the loaded package hashes with the intended deployed package;
- [x] inspect the complete workbook-resolution and role-workbook provisioning
  paths; and
- [x] identify whether the defects are isolated mistakes or symptoms of a
  broader launcher/context design problem.

Refactoring is allowed when the assessment and protecting RED demonstrate that
it is the safer or more maintainable correction. A justified refactor may
consolidate duplicated launcher/context logic, split responsibilities, or
change the affected form/controller boundary. It must:

- [x] preserve the normative D12 package, Domain authority, event,
  `System_Key`, extensible-header, and captured-workbook contracts;
- [x] state why a local patch is insufficient;
- [x] define the bounded files, procedures, and contracts being changed before
  implementation;
- [x] add focused tests for every contract moved or generalized;
- [x] remain reversible and preserve plan 020 regression evidence;
- [x] never accept an arbitrary active workbook, canonical runtime workbook, or
  XLAM as operator authority;
- [x] never use a non-test operational NAS warehouse for write-capable tests;
  and
- [x] never build or deploy XLAMs while Excel has the relevant add-ins or
  workbooks open.

## 5. D13 execution sequence

### Entry gate — identify the deployed runtime without mutation

Use the read-only runtime extractor and targeted source inspection to record:

- [x] the full paths and hashes of the five loaded XLAMs;
- [x] whether those hashes match the approved `deploy/current` manifest;
- [x] the selected warehouse ID, station ID, target type, and redacted UNC/NAS
  runtime root;
- [x] the active/open workbook classifications, without row-level operational
  values;
- [x] whether an eligible station-local role workbook already exists for each
  launcher;
- [x] the exact public callbacks:
  `modTS_Received.ShowReceivingForm`,
  `mProduction.BtnOpenProductionForm`, and
  `modTS_Shipments.BtnOpenShipmentsForm`; and
- [x] before/after hashes proving the read-only capture changed no inspected
  workbook.

This gate must not save, refresh, repair, process, close, or create any
operational workbook.

### Slice 0 — packaged launcher RED and stage diagnostics

Create a focused packaged test that invokes the same three public callbacks as
the Operations ribbon under these states:

- [x] connected and signed in, with no eligible role operator workbook open;
- [x] a config/auth/snapshot/canonical workbook active;
- [x] a saved eligible role workbook active;
- [x] an unrelated ordinary workbook active; and
- [x] a captured role workbook closed after form creation.

Record meaningful RED for:

- [x] Receiving refusing to provision/open its station-local workbook;
- [x] Production returning the observed type mismatch; and
- [x] Shipping returning the observed type mismatch.

Add only enough safe diagnostic structure to report the failing launcher stage,
`Err.Number`, sanitized `Err.Source`, and `Err.Description`. A compile failure,
missing workbook, unavailable NAS, or broken harness is not RED.

Gate:

- [x] all three failures reproduce through the packaged callbacks;
- [x] the harness distinguishes resolution, provisioning, surface repair, form
  initialization, render, and modeless-show stages;
- [x] no canonical or unrelated workbook is selected as operator authority; and
- [x] the RED report contains no secrets or row-level operational data.

### Slice 1 — Receiving station-local workbook self-provisioning

Implement a narrow Receiving launcher path that:

- [x] first reuses an already open eligible Receiving operator workbook;
- [x] otherwise opens the existing station-local Receiving operator workbook
  resolved from the current warehouse/station context;
- [x] otherwise creates and saves that workbook through a declared Core-owned
  primitive boundary;
- [x] creates the managed Receiving and inventory read-model surfaces with no
  `ROW` header;
- [x] refreshes from the selected warehouse snapshot without making the
  operator workbook authoritative;
- [x] binds the modeless Receiving form to the created/opened workbook;
- [x] activates or presents the operator workbook without changing the selected
  NAS target or signed-in invSys user;
- [x] never saves or mutates canonical config, auth, inventory, snapshot,
  inbox, outbox, or Domain workbooks; and
- [x] returns a precise recovery message if creation/opening cannot complete.

The path must be idempotent: repeated Receiving-control clicks reuse the same
eligible saved workbook and valid form instance rather than creating duplicate
files, forms, or event subscriptions.

If a user closes the modeless Receiving form while leaving its operator
workbook open, the next Receiving-control click must discard the disappeared
form reference, recreate the form, and preserve the same captured workbook.

Gate:

- [x] meaningful packaged RED/GREEN is recorded;
- [x] the actual Receiving ribbon callback creates/opens the workbook;
- [x] the workbook is station-local and saved as a supported macro-enabled
  format;
- [x] the form remains bound to its captured workbook when another workbook is
  activated;
- [x] closing only the Receiving form and clicking Receiving again recreates
  one valid form against the same workbook without an automation error;
- [x] custom local columns survive reopening and refresh;
- [x] no canonical NAS workbook hash changes merely from opening the form; and
- [x] plan 020 Receiving, packaged XLAM, RibbonX, static-tooling, and
  maintenance regressions remain GREEN.

### Slice 2 — Production launcher type-mismatch repair

Use Slice 0 stage evidence to repair only the failing Production boundary.

Required behavior:

- [x] `mProduction.BtnOpenProductionForm` validates the target before surface
  repair or form initialization;
- [x] the callback never falls back to `ThisWorkbook` or another XLAM as
  operator authority;
- [x] config, auth, inventory, snapshot, inbox, outbox, and Domain workbooks
  are rejected as operator workbooks;
- [x] surface repair receives the declared primitive/workbook type at the
  correct project boundary;
- [x] `frmProduction` initializes and opens modelessly against one captured
  eligible workbook; and
- [x] a missing eligible workbook is created or opened station-locally through
  the declared Core-owned role-workbook primitive, then captured by the form.

Gate:

- [x] focused packaged RED/GREEN identifies and removes the exact mismatch;
- [x] saved-workbook initialize, captured-workbook, two-batch, restart, and
  layout regressions remain GREEN;
- [x] no new same-project `Application.Run` or broad error suppression is
  introduced; and
- [x] static bloat/dynamic-call metrics do not regress without a reviewed
  exception.

### Slice 3 — Shipping launcher type-mismatch repair

Use Slice 0 stage evidence to repair only the failing Shipping boundary.

Required behavior:

- [x] `modTS_Shipments.BtnOpenShipmentsForm` validates an eligible Shipping
  operator workbook before quiet UI, surface repair, clipboard handling, form
  initialization, or autosync registration;
- [x] config, auth, inventory, snapshot, inbox, outbox, and XLAM workbooks are
  rejected as operator authority;
- [x] `Nothing` is never passed to a typed workbook operation;
- [x] `frmShipmentsTally` initializes and opens modelessly against one captured
  eligible workbook; and
- [x] a missing eligible workbook is created or opened station-locally through
  the declared Core-owned role-workbook primitive, then captured by the form.

Gate:

- [x] focused packaged RED/GREEN identifies and removes the exact mismatch;
- [x] Boxing tabs, Shipments staging, exact lock release, Shipments Sent,
  replay, restart, and overlay regressions remain GREEN;
- [x] no canonical NAS workbook is polluted or selected as the operator
  workbook; and
- [x] static bloat/dynamic-call metrics do not regress without a reviewed
  exception.

### Slice 4 — deployed-package and dedicated NAS test-runtime acceptance

After isolated RED/GREEN and regressions pass:

- [x] rebuild the five-package set with Excel closed;
- [x] record manifest paths and hashes;
- [x] install/load the approved package set through the real account-scoped
  deployment path;
- [x] use a dedicated generated test warehouse on the actual NAS/UNC path;
- [ ] keep station-local operator workbooks local to the user profile;
- [ ] connect, sign in, and invoke Receiving, Production, and Shipping through
  the visible Operations ribbon controls;
- [x] restart Excel and repeat the launcher checks; and
- [x] compare before/after hashes so only the dedicated test runtime and
  intended station-local operator workbooks changed.

Do not point this acceptance run at a warehouse containing non-test operational
inventory.

### Slice 4a â€” Admin Seed Demo Inventory acceptance blocker

The first downstream Admin control checkpoint on 2026-08-02 found that
**Seed Demo Inventory** flashed Excel and ended with `Application-defined or
object-defined error`. This is a newly discovered Release 1 acceptance blocker,
not a reopening of the GREEN Operations launcher contracts.

Required behavior:

- [x] the packaged test invokes `modAdmin.Seed_DemoInventory`, not only
  `modAdminConsole.SeedDemoInventoryForAutomation`;
- [x] a valid current warehouse target is used directly without resolving a
  canonical Config/Auth/inventory workbook as an Admin operating surface;
- [x] general View Warehouses scanning remains available outside the seed
  callback;
- [x] failures identify context resolution versus queue/processor application
  with error number, sanitized source, and description;
- [x] the seed creates three new collision-free `System_Key` values with
  `Condition=GOOD` and no managed `ROW` header; and
- [x] Config remains byte-for-byte unchanged, Auth table data remains
  unchanged, and the intended dedicated test inventory/inbox runtime changes.

Gate:

- [x] meaningful packaged RED records the callback failing to reach/complete
  the selection path within 45 seconds;
- [x] packaged callback GREEN is recorded against an isolated generated test
  warehouse with a canonical Config workbook active and read-only;
- [x] Create Warehouse / repeated seed D14 lifecycle remains GREEN at 15/15;
- [x] packaged XLAM validation remains GREEN at 54/54;
- [x] packaged RibbonX remains GREEN at 136/136;
- [x] the ordered Release 1 full chain remains GREEN at 30/30; and
- [x] the operator confirms the visible Admin **Seed Demo Inventory** control
  succeeds in the dedicated NAS test warehouse.

### Slice 4b — Seed snapshot and operator-read-model round trip

The 2026-08-04 operator checkpoint proved that the corrected public callback
returns and the processor reports one applied seed event, but the reported
refresh did not visibly prove the D14 round trip. Later read-only inspection
found the demo entities in all three saved role projections. The expanded
packaged test then proved the exact callback-to-Receiving round trip. The
original visible checkpoint remains failed until repeated on the specified
Receiving inventory controls.

Required behavior:

- [x] the same public `modAdmin.Seed_DemoInventory` callback queues and applies
  one batch event containing three new durable entities;
- [x] the canonical event/entity projections contain exactly three new unique
  `System_Key` values with `Condition=GOOD`;
- [x] the processor publishes a snapshot containing those three entities;
- [x] a captured saved role operator workbook refreshes from that snapshot and
  exposes the same three entities; and
- [x] acceptance evidence distinguishes event application, snapshot
  publication, operator refresh, and Receiving-list visibility instead of
  treating `Applied=1` alone as end-to-end proof.

Gate:

- [x] extend the packaged public-callback harness through snapshot inspection
  and a captured operator-workbook refresh;
- [x] preserve the operator-visible RED while the callback reported an applied
  event;
- [x] obtain GREEN for canonical, snapshot, operator counts/keys, and the
  Receiving Refresh form action; and
- [x] repeat the visible dedicated-NAS seed/refresh checkpoint.

### Slice 4c — Production native-window resize acceptance blocker

The 2026-08-04 Production screenshot shows the native form window maximized
while the MultiPage and child controls retain a small base-size footprint in
the upper-left corner. Existing Slice 9 evidence proves that the window can
enter the maximized state and that manually assigned form sizes have no
overlaps. It does not prove that a native maximize resizes the actual UserForm
client area and anchored content.

Required behavior:

- [x] the packaged test enters through
  `mProduction.BtnOpenProductionForm` and one saved Production operator
  workbook;
- [x] native maximize/restore changes the measurable UserForm client/content
  geometry, not only the outer HWND state;
- [x] `mpProduction`, the status box, Close button, and each active page use
  the available client area without clipping, overlap, or an unused blank
  majority; and
- [x] restore returns to a readable supported size.

Gate:

- [x] record focused native-maximize RED from the same public launcher path;
- [x] obtain automated bounds/fill/overlap GREEN for all four pages;
- [x] capture redacted before/after screenshots; and
- [ ] repeat visible Production maximize/restore acceptance.

### Slice 4d — Shipping `ROW` display-key conflict

The form-control inventory found that `frmShipmentsTally` still creates
`lblRow`, `txtRow`, `hdrShipRow`, and generated shipment-line `ROW` headers.
This contradicts v4.11 D14, which prohibits `ROW` as a managed runtime header,
display key, compatibility field, or authority path.

Gate:

- [x] add focused packaged form-action RED that proves the visible Shipping
  selection/staging path still exposes or depends on `ROW`;
- [x] replace the form and backing path with exact `System_Key` identity;
- [x] retain Shipping/Boxing lock, stage, hold/return, Shipments Sent, restart,
  and version regressions; and
- [x] regenerate static maintenance evidence with no prohibited managed
  `ROW` control/header.

Evidence: `tests/integration/plan022_slice4_controls_results.md`. The focused
contract is 12/12 GREEN, the packaged Shipping form-action proof preserves exact
string `System_Key`, and the static baseline records the bounded oversized-form
and controller exceptions without dynamic-call or duplicate-body regression.

### Slice 4e — Release 1 surface cleanup and Inventory Viewer

The 2026-08-08 control review approved the current Operations and Admin
workflows for Release 1, approved controls being visible to every signed-in
user, and directed removal of nonfunctional shells. The same review identified
one missing Release 1 surface: a read-only, at-a-glance inventory viewer that
does not require opening a role form. D4 is corrected first because the runtime
already uses one shared Core item-search form while four role-named source forms
and dynamic-template copies are empty shells.

Required behavior:

- [x] package only the reviewed active Admin, Operations, Core, Receiving,
  Production, and Shipping forms; remove empty or unreachable form shells after
  static reachability and packaged regression evidence;
- [x] retain the Purchasing stub in the active Receiving form;
- [x] expose **Inventory Viewer** on Operations to every signed-in user;
- [x] open one reusable, modeless, resizable Viewer bound to the selected
  warehouse and show item code, item, UOM, quantity, location, and condition;
- [x] load the Viewer only from the published inventory snapshot on explicit
  refresh and leave every inspected workbook byte-for-byte unchanged; and
- [x] preserve role-aware item search through the single Core
  `frmItemSearch` runtime form after obsolete role-named forms are removed.

Gate:

- [x] focused source/package RED records the obsolete shells and absent Viewer;
- [x] the packaged public Viewer callback, repeated launch, filtering, and
  read-only snapshot proof are GREEN;
- [x] Receiving, Production, Shipping, Admin, item-search, RibbonX, collision,
  and full-chain regressions remain GREEN; and
- [x] static maintenance evidence records the reviewed form set with no removed
  component still reachable or packaged.

### Slice 4f — Complete Receiving-to-Production-to-Shipping demo inventory

The three-row bootstrap proved event, snapshot, and role-projection transport,
but it does not provide enough material classes for an operator to exercise a
complete Release 1 workflow. The reviewed sanitized demo catalog supplies the
Release 1 test fixture; its legacy local-table mutation and `ROW` identity path
are not reused.

Required behavior:

- [x] the public Admin **Seed Demo Inventory** action queues one batch event
  containing 19 new durable entities: raw ingredients, WIP, finished/shippable
  goods, cases/boxes, and tins;
- [x] every entity receives a new unique immutable `System_Key` and
  `Condition=GOOD`, with item name, UOM, location, description, and category
  metadata sufficient for the role projections;
- [x] one applied event publishes all 19 entities to the snapshot and a
  Receiving refresh exposes all 19; and
- [x] this slice originally treated repeated seed as an intentional new-entity
  action; Slice 4n supersedes that behavior with active-group idempotence.

Gate:

- [x] extend the existing packaged public-callback harness and record a
  meaningful 3-versus-19 RED;
- [x] obtain GREEN for event, canonical, snapshot, operator, Receiving form,
  identity, condition, and category coverage; and
- [x] preserve unchanged Config/Auth evidence and the full D14 lifecycle suite.

### Slice 4g — Shipping fixed status/message anchor

The 2026-08-08 Shipping screenshots prove that height resize currently moves
the status/message control below Search Boxes. The status area must remain at
its established top position and may stretch horizontally, while the inventory
list and lower content consume added height.

Gate:

- [x] a focused test through `BtnOpenShipmentsForm` records that status Top
  changes after height resize;
- [x] anchor the status/message control Left + Top + Right with fixed height;
- [x] prove it remains above Search Boxes through grow, shrink, maximize, and
  restore; and
- [x] retain Shipping launcher, list, Boxing, lock, restart, and resize
  regressions.

### Slice 4h — Boxing form usability and layout

The 2026-08-09 visible checkpoint accepted the general form-resize behavior but
found that Box Builder and Box Maker were omitted from the anchor registration.
The same checkpoint found the Builder component inventory too narrow, without
a search control or reliable column/header alignment. Inventory items without
a Shipping BOM version must display `NA`; version labels belong only to the
versioned Shipping BOM/design relationship.

Required behavior:

- [x] the Inventory Viewer ribbon button uses a visible built-in Excel icon;
- [x] Box Builder and Box Maker list controls resize with the native form while
  keeping buttons, fields, and lists readable and non-overlapping;
- [x] Box Builder exposes a component-inventory search that filters the already
  loaded managed inventory choices without mutating an authority workbook;
- [x] Box Builder and Box Maker list headers use the same declared widths and
  resize/move contract as their list columns;
- [x] non-versioned managed inventory displays `NA` in a Version column, while
  versioned Shipping BOM packages display their active version label; and
- [x] the public Shipping launcher remains the packaged entry boundary.

Gate:

- [x] focused public-launcher RED records omitted Boxing anchors, absent search,
  missing/blank non-versioned display, and header drift;
- [x] packaged grow/shrink/maximize/restore and component-search proof is GREEN;
- [x] Box Builder save/version, Box Maker make/unmake, Shipping, and restart
  regressions remain GREEN; and
- [ ] the user confirms the two Boxing tabs remain readable after resize.

### Slice 4i — Receiving Entries History

The 2026-08-09 visible checkpoint found that the Receiving form's top
per-entity inventory list duplicates the at-a-glance Viewer and appears
duplicated because several durable entities may share one item code. The top
list is therefore a Receiving Entries History projection, not a second current
inventory viewer. Selecting what to receive remains a separate, deduplicated
managed-item control backed by the captured operator read model.

Required behavior:

- [x] the top list is titled **Receiving Entries History** and loads completed
  entries from the captured workbook's `ReceivedLog` table;
- [x] history search filters receipt date/user/reference/item/code/location and
  does not mutate staging or authority workbooks;
- [x] a separate deduplicated managed-item selector supplies the item code and
  source `System_Key` needed to stage a new receipt entity;
- [x] each staged receipt still generates a new immutable `System_Key` before
  queueing, while history retains `EventId` and `System_Key` logging identity;
  and
- [x] Refresh updates history, the managed-item selector, and staging without
  using the history row as inventory authority.

Gate:

- [x] focused RED proves the top list currently reloads `invSys` rather than
  `ReceivedLog` and that item selection is not separated from history;
- [x] the packaged `ShowReceivingForm` Refresh and Confirm Writes paths are
  GREEN with captured-workbook history and staging preserved; and
- [x] Receiving D14, snapshot refresh, reopen, and full-chain regressions remain
  GREEN.

### Slice 4j — Full workflow control readiness

The 2026-08-16 control review deliberately expands Slice 4 after the prior
Boxing layout was visibly rejected. It preserves the D12 Operations package and
D13 test-first boundary while making the current controls sufficient for a
complete Receiving -> Production Run - List -> Box Maker -> Shipping operator
test. Production Run - Tree remains experimental and is not changed by this
slice.

Required behavior:

- [x] the saved role workbook names remain exactly
  `<WarehouseId>.Receiving.Operator.xlsm`,
  `<WarehouseId>.Production.Operator.xlsm`, and
  `<WarehouseId>.Shipping.Operator.xlsm`;
- [x] ordinary station identity is derived from the Windows computer name in
  warehouse creation, connection, seed, runtime-workbook, Admin, and ribbon
  fallback paths; the connection form no longer asks the user to select a
  bespoke station or station-inbox checkbox;
- [x] Admin exposes one **Design Lifecycle** ribbon launcher, with Release and
  Obsolete retained as actions inside the lifecycle form;
- [x] **Test Environment Setup** remains an Admin-only isolated fixture
  provisioner and cleanup utility, clearly distinguished from ordinary station
  identity and operator work;
- [x] the Add/Edit Inventory mode selector reads **Add Item Mode**;
- [x] Seed Demo Inventory contains 24 new `System_Key` entities, including
  shipping carton, divider, label, tape, and void-fill consumables needed for
  Box Designer and Box Maker tests;
- [x] the visible **Box Designer** and Box Maker pages use explicit full-width
  vertical layouts, recalculated column headers, and non-overlap geometry after
  form resize;
- [x] visible box wording uses **Alternative** rather than software-release
  **Version**; compatible stored BOM keys such as `v1` remain internal data;
- [x] Box Designer choices, saved BOM components, Box Maker availability, and
  Make/Unmake payloads preserve exact string `System_Key` identity end to end;
- [x] Receiving has a dedicated searchable item-results list, requires a
  receiving Location, accepts optional Lot number, and carries both through
  tally, aggregate, event note, and receipt history surfaces; and
- [x] Production Run - List accepts and applies a batch scale from `0.001%`
  through `1000%`, defaulting to `100%`.

Gate:

- [x] focused source RED recorded 1 pass / 12 failures before implementation;
- [x] focused source GREEN is 18/18, with Demo Seed 4/4, R1 final controls
  12/12, Slice 5 behavior locks 13/13, Slice 10 Receiving 10/10, Slice 11
  Shipping/Boxing 11/11, Production layout 8/8, and ribbon generation 46/46;
- [x] static source inventory keeps literal `Application.Run` at 9, unresolved
  dynamic calls at 48, and improves duplicate-body candidates from 185 to
  184. The deterministic baseline is 150 components and 4,614 procedures. The
  reviewed Slice 4j feature exception permits required growth in the existing
  oversized surfaces and services: `frmShipmentsTally` (+260 lines),
  `frmProduction` (+145), `frmReceiving` (+123),
  `modRoleWorkbookSurfaces` (+103), `modTS_Shipments` (+85),
  `modTS_Received` (+39), `modReceivingPostingService` (+25),
  `modBoxingService` (+17), `mProduction` (+16),
  `modAdminInventorySeed` (+10), and `frmSeedInventory` (+2). Those additions implement the requested operator
  layouts, header alignment, exact `System_Key` BOM/event identity, receiving
  attributes, seed catalog, and batch-scale contract. The one new component is
  a bounded 10-line shared computer-station identity module, not a parallel
  workflow implementation;
- [x] rebuild Core, Inventory Domain, Designs Domain, Operations, and Admin only
  after every relevant Excel workbook/add-in is closed;
- [x] packaged XLAM validation is 74/74 and Ribbon validation is 142/142;
  packaged Production scaling, Receiving durability/search/header alignment,
  and Shipping status/Box Designer/Box Maker resize and identity probes each
  pass their public callback contract;
- [x] the isolated public Admin callback publishes all 24 unique `System_Key`
  entities with `Condition=GOOD`, and the ordered isolated Release 1 full-chain
  validator is GREEN at 30/30 through restart and reconciliation;
- [ ] repeat the 24-row Admin seed validator and Release 1 full workflow on the
  dedicated NAS warehouse; and
- [ ] obtain visible user confirmation for Box Designer/Box Maker resize,
  Receiving search/location/lot, Production List scaling, and the complete
  seed-to-ship workflow.

### Slice 4k — Selected warehouse/session binding

The 2026-08-17 visible checkpoint showed that **Connect Server** reached the NAS
but both target choices were rejected because `tblStationConfig` still held the
legacy station `S1`. The ribbon therefore retained `invsys_Zenbook_WH`, the
sign-in form displayed `<roaming>`, and the user's WHT credential was checked
against the wrong Auth workbook. The failure dialog compounded the problem by
showing the prior connection status rather than the failed attempt status.

Required behavior:

- [x] the exact **Send To** public action accepts the automatically discovered
  Windows computer name and binds the selected warehouse before sign-in;
- [x] when that computer is missing from a valid warehouse config, the system
  safely enrolls only the exact current computer station; arbitrary or stale
  station identifiers remain rejected;
- [x] target selection validates before atomically replacing the current
  warehouse, station, and runtime root;
- [x] changing any target identity signs out the prior invSys session, so no
  authenticated state crosses warehouse boundaries;
- [x] failure dialogs report the attempted selection result and say that the
  prior target remains current; and
- [x] NAS automation uses the Windows computer name rather than hard-coded
  `S1`, and read-only runtime extraction handles one-workbook scalar results
  without mutating an inspected workbook.

Gate:

- [x] focused behavioral RED is recorded at 0/1 after the valid Excel fixture
  reaches the legacy-station rejection;
- [x] focused current-computer selection and same-action ribbon tests are 2/2;
- [x] Core target/auth/write regression is 28/28;
- [x] packaged XLAM and Ribbon validation remain 74/74 and 142/142;
- [x] dedicated NAS validation is 12/12 across two clean Excel sessions for
  `WHT7025AE` / `X1-PRO-AI`, including selected-target, launcher, package-hash,
  and runtime read-only-safety checks;
- [x] the isolated ordered Release 1 workflow remains 30/30 through restart and
  reconciliation, Tool B remains 62/62, and the deterministic static baseline
  remains 19/19; and
- [ ] the user visibly confirms that selecting `WHT7025AE` changes the ribbon
  target, the sign-in form shows `WHT7025AE` / `X1-PRO-AI`, and the existing
  warehouse-specific credential signs in.

### Slice 4l — Ribbon session-state controls

The 2026-08-17 visible checkpoint confirmed target selection but showed that
the deployed Operations server label stayed on the prior warehouse until
Runtime Context caused a later refresh. It also showed ambiguous generic
Sign In/Sign Out wording and no server-disconnect action. The root cause of the
stale label was exact: warehouse selection invalidated retired role-specific
Ribbon IDs but omitted `ddOperationsWarehouseTarget`,
`lblOperationsServerStatus`, and `lblOperationsAccessStatus`.

Normative clarification:

- **invSys Sign In / invSys Sign Out** controls only the invSys identity and
  capability session; invSys Sign Out retains server access for user switching;
- **Server Sign In / Server Sign Out** controls the NAS/Windows SMB layer;
  Server Sign Out also clears invSys authentication and the selected warehouse;
- disconnected invSys Sign In fails closed with an instruction to use Server
  Sign In and must not revive remembered target state; and
- all capability-gated operator actions remain disabled until server sign-in,
  target selection, and invSys sign-in are complete.

Gate:

- [x] focused reachable-source RED is 0/8 for the reported ribbon/session
  behaviors, followed by an 8/9 RED proving disconnected Operations still
  displayed `Access: Ready`;
- [x] focused GREEN is 9/9 and same-action Excel tests are 2/2;
- [x] Core target/auth/session regression is 30/30 and ribbon generation is
  48/48;
- [x] packaged XLAM and Ribbon validation are 74/74 and 140/140;
- [x] dedicated `WHT7025AE` / `X1-PRO-AI` NAS validation is 16/16 across two
  clean sessions, including immediate selected-target callback state, a real
  SMB disconnect, and successful reconnect;
- [x] the isolated ordered Release 1 workflow remains 30/30, Plan 022 launcher
  contracts remain 24/24, Slice 4j remains 18/18, and the deterministic static
  baseline remains 19/19; and
- [ ] the user visibly confirms immediate Send To status refresh, both explicit
  toggle labels, disabled operator controls after Server Sign Out, and the
  connect-first prompt from disconnected invSys Sign In.

### Slice 4m — current-computer Admin capability transition

The 2026-08-17 visible checkpoint reached the correct `WHT7025AE` Auth
workbook and displayed station `X1-PRO-AI`, but the Admin invSys sign-in form
returned **User lacks the required capability**. The credential had passed;
the remaining failure was exact station scope: the generated warehouse still
held the user's `ADMIN_MAINT` grant at the obsolete `S1` placeholder while the
selected target correctly used the Windows computer name.

Normative clarification:

- a valid invSys credential, not Server Sign In or target selection, is the
  authority boundary for a one-time legacy station transition;
- only the same user's effective active `S1` capabilities may be copied to the
  exact current-computer station;
- validity dates, warehouse scope, existing current-station rows, and denies
  are preserved; and
- a user who lacked the requested legacy capability remains denied and gains
  no current-station capability.

Gate:

- [x] same-path behavioral RED is 1/2: Admin migration fails while the
  missing-capability security case passes;
- [x] focused GREEN is 3/3, including explicit current-station deny
  preservation, and the complete Core NAS/target/auth/session block is 33/33;
- [x] rebuilt packaged XLAM and Ribbon validation are 74/74 and 140/140;
- [x] the isolated ordered Release 1 workflow is 30/30, Plan 022 launcher
  contracts are 24/24, Slice 4j is 18/18, and Slice 4l is 9/9;
- [x] dedicated `WHT7025AE` NAS validation is 16/16 using the automation user
  without changing the human UAT credential;
- [x] regenerated deterministic static evidence is 19/19 with 150 components,
  4,633 procedures, 8 literal `Application.Run` targets, 47 unresolved dynamic
  calls, and 184 duplicate-body candidates; and
- [ ] the user visibly signs into Admin as the existing warehouse user and
  confirms the Admin ribbon controls enable at `WHT7025AE` / `X1-PRO-AI`.

### Slice 4n — selectable demo-inventory lifecycle and operator projections

The 2026-08-17 Receiving checkpoint showed duplicate-looking search rows after
repeated demo seeds. The operator also needs to choose the fixture used for a
seed and to remove demo state without erasing canonical audit history.

Required behavior:

- [x] the Admin ribbon mirrors Operations with a Core-owned live **Send To**
  warehouse selector and shared target-selection callback;
- [x] the Admin ribbon opens one **Demo Inventory** form with explicit **Seed
  Demo Inventory**, **Delete Demo Inventory**, **Upload Data Set**, and **Delete
  Data Set** actions;
- [x] the form selects the immutable built-in Release 1 workflow kit or a
  validated CSV stored in the selected warehouse data-set library; Upload
  imports a definition and Seed applies the selected definition;
- [x] Delete Data Set removes only a selected uploaded definition, never the R1
  kit or inventory already seeded from that definition;
- [x] repeated built-in or uploaded seeds skip active item/location/condition
  groups, while missing or fully depleted groups receive new immutable keys;
- [x] confirmed Delete posts exact-`System_Key` depletion adjustments for every
  active `DEMO-` entity and retains canonical entity/event history;
- [x] uploaded CSV validation is all-or-nothing and requires `DEMO-` item codes,
  positive quantity, complete required fields, and no duplicate active group;
  and
- [x] Inventory Viewer and Receiving choice projections aggregate matching
  active entities and omit nonpositive totals.

Gate:

- [x] focused RED records absent lifecycle actions, repeat-seed duplication,
  stale zero-quantity projections, and absent selected-data-set routing;
- [x] packaged public callback GREEN covers built-in selection, persistent CSV
  import/selection, repeated seed, exact-key inventory depletion, uploaded
  definition deletion with seeded stock retained, R1 protection, invalid-file
  rejection, and canceled deletion;
- [x] Receiving aggregation, Generate Warehouse/Create Warehouse, packaged
  XLAM, RibbonX, and the ordered Release 1 chain remain GREEN; and
- [ ] the user visibly confirms the revised form and chosen data set against the
  dedicated NAS test warehouse.

### Slice 4o — Receiving condition, inbound returns, and aggregate repair

The 2026-08-19 visible Receiving checkpoint confirmed Demo Inventory seed and
delete, then exposed three operator-contract gaps: closing Demo Inventory emits
a misleading cancellation dialog, receipt Condition cannot be established by
the receiver, and Aggregate Received may remain a stale one-row projection
while Received Tally contains several lines. The user also requires inbound
returned goods to be received explicitly rather than hidden inside ordinary
receipt wording.

Required behavior:

- [x] remove the redundant Demo Inventory **Cancel** button; closing the form
  without choosing an action is silent and does not report that a completed
  seed or delete was cancelled;
- [x] establish editable line-level `Condition` in Receiving, defaulting to
  `GOOD`, and carry it through tally, aggregate, event, canonical inventory,
  and receipt history;
- [x] keep Inventory Viewer read-only and keep `Lot` independent from
  `Condition` and `System_Key`;
- [x] rebuild Aggregate Received from every staged line, grouping by receipt
  type/reference/item/location/lot/condition/return reason and summing repeated
  matches;
- [x] add a functional Receiving **Returns** page for inbound returned goods,
  capturing return reference, reason, quantity, location, optional lot, and
  condition, and creating a new immutable `System_Key` through the existing
  `RECEIVE` event boundary; and
- [x] preserve the Purchasing stub and the ordinary Receiving workflow.

Gate:

- [x] focused same-handler RED records the current Cancel control, absent
  Condition/Returns contract, and incomplete aggregate rebuild;
- [x] focused GREEN exercises the Demo Inventory form contract and the public
  Receiving add/refresh/confirm actions against a generated operator workbook;
- [x] packaged XLAM/Ribbon, D14 identity, Receiving durability, and the ordered
  Release 1 chain remain GREEN; and
- [ ] the user visibly confirms mixed-condition receipt lines, aggregate totals,
  and one inbound return against `WHT7025AE`.

Automated evidence recorded 2026-08-19: focused source contract 5/5 after its
0/5 RED; same-form public-action GREEN 4/4 after its 0/2 RED; packaged Slice 4o
actions 5/5; Receiving workbook-surface regressions 28/28 over the affected
bootstrap/migration range; packaged XLAM 74/74; packaged RibbonX 142/142; and
the ordered deployed Release 1 chain 30/30. Static maintenance evidence is
deterministic 19/19 with 150 components, 4,680 procedures, 961 scanner
candidates, 963 reviewed candidates, and 24 oversized-module ratchets.

### Slice 4p — Receiving aggregate semantics, Return labels, and save batching

The 2026-08-19 visible Slice 4o checkpoint proved the aggregate projection is
complete, then identified that it no longer performs its established tally
role: matching items from different PO/BOL references appear on separate rows.
The same checkpoint showed that Returns retains Receiving titles, omits
Condition from Return Item Results, and produces repeated Excel Saving notices
(four during sign-in, more than sixty during Confirm Writes, and thirty-three
during Confirm Returns).

Required behavior:

- [x] Aggregate Received/Returns groups matching receipt lines by receipt type,
  item code, UOM, location, lot, and Condition, sums quantity, and concatenates
  distinct references in first-seen order; different Conditions remain on
  separate rows;
- [x] `ReceivedTally`, not the aggregate projection, remains the separately
  keyed queue/log authority so aggregation cannot collapse `System_Key` or
  `EventId`;
- [x] Receiving/Return item results display Condition;
- [x] Returns mode titles its projections **Return Entries History**, **Return
  Tally**, and **Aggregate Returns** while Receiving mode retains its Receiving
  titles; and
- [x] healthy sign-in does not dirty/save unchanged Config/Auth workbooks, and
  multi-row confirm persists once per safe workbook/artifact phase rather than
  once per row.

Gate:

- [x] same-handler RED proves reference-separated aggregates, missing Return
  Condition/title behavior, read-load workbook dirtiness, and per-row processor
  persistence;
- [x] focused GREEN proves concatenated references/totals with Condition
  separation and distinct staged identities through Confirm;
- [x] packaged form-action and processor evidence reports bounded save cycles
  independent of receipt-row count;
- [x] packaged XLAM/Ribbon, D14 identity/durability, ordered Release 1 chain,
  and static ratchets remain GREEN; and
- [ ] the user visibly confirms the revised aggregate, Return-mode UI, and
  bounded Saving notifications against `WHT7025AE`.

Automated evidence recorded 2026-08-19: behavior locks 289-294 are GREEN 6/6
after the four newly introduced contracts were RED; focused Slice 4p static
contract 6/6; packaged Slice 4o/4p form actions 5/5; packaged XLAM 74/74;
packaged RibbonX 142/142; live role workflow 46/46; and the standalone ordered
Release 1 chain 30/30. The three-event Receiving processor proof reports exactly
three persistence saves: canonical inventory, one batched outbox append, and one
inbox-status save. Deterministic maintenance evidence remains GREEN with 150
components, 4,686 procedures, and 964 scanner candidates.

### Slice 4q — outbound Return/Dump inventory disposition

The 2026-08-19 visible Slice 4p checkpoint proved the revised labels and
Condition column, then exposed that the underlying Returns contract was
backwards: returning 50 units of an existing 100-unit DAMAGED entity created a
new inbound entity and displayed 150. The operator requires both vendor return
and trash disposal. This slice deliberately supersedes the inbound-return
behavior recorded in Slice 4o; the normative D14 rule now defines outbound
disposition.

Required behavior:

- [x] Returns exposes a required **Disposition** selector with `RETURN` and
  `DUMP`, defaulting to `RETURN`;
- [x] both actions accept a positive operator quantity but reduce canonical
  and Viewer on-hand quantity by that amount through distinct `RETURN`/`DUMP`
  audit event types;
- [x] disposition targets existing exact `System_Key` entities, creates no new
  inventory key, preserves item/location/lot/Condition, and rejects overdraw;
- [x] an aggregate search choice spanning several entities is deterministically
  allocated into separately keyed staged/event rows without crossing
  item/location/Condition boundaries;
- [x] Return Tally, Aggregate Returns, and Return Entries History show the
  selected disposition and action quantity clearly; and
- [x] the `RECEIVE_POST` authorization boundary, batched persistence, ordinary
  Receiving workflow, Inventory Viewer read-only behavior, and D12 package
  boundary remain intact.

Gate:

- [x] same-handler RED proves the current Return action creates inventory,
  creates a new key, and lacks the required selector;
- [x] focused Domain/processor RED proves `RETURN`/`DUMP` are unsupported and
  overdraw/exact-key rules are absent;
- [x] focused GREEN proves 100 minus a 50-unit RETURN equals 50, RETURN plus
  DUMP depletion remains exact-key and condition-separated, and no new entity
  is created;
- [x] packaged Returns form action, Viewer refresh, XLAM/Ribbon, full workflow,
  restart, and static maintenance regressions remain GREEN; and
- [x] the user visibly confirms Return/Dump selection and resulting quantities
  against `WHT7025AE`.

Automated evidence on 2026-08-19: the focused static contract began at 0/6
and the same-handler and Domain tests both returned 0 before implementation.
The aligned VBA range is now 7/7, the focused static contract 6/6, and the
packaged Returns action 5/5. The standalone Viewer proof loaded/filter-reused
its snapshot without mutation; packaged XLAM is 74/74, RibbonX 142/142, live
role workflow 46/46, and the clean ordered Release 1 chain 30/30. Static
maintenance is deterministic 19/19 with 150 components, 4,698 procedures,
965 scanner candidates, and a narrow protected Slice 4q procedure ceiling;
the reviewed cleanup contract remains 11/11 with no component, candidate,
duplicate-body, or late-binding regression.

The first visible Slice 4q retest exposed an Excel table-placeholder blocker:
after a valid DAMAGED item was selected, `ReceivedTally` could retain a second
wholly blank physical row and aggregation rejected it as a business record with
`ReceivedTally contains a blank System_Key.` The same public Returns form
action reproduced this as RED (`Returned=0`). Aggregation and confirmation now
remove only rows whose identity, item, reference, event, and quantity fields
are all blank; any partially populated row still reaches normal validation and
cannot be silently discarded. The focused same-handler test is GREEN, the
aligned range remains 7/7, packaged Returns remains 5/5 with the extra blank
row present, packaged XLAM is 74/74, live workflows 46/46, the clean ordered
chain 30/30, and static maintenance remains 19/19 plus 11/11.

The next visible retest selected a valid 50-unit DAMAGED entity and requested a
5-unit `RETURN`, but **Add Disposition** failed with Excel error 1004 and only
the generic text `Application-defined or object-defined error`. Read-only
diagnostics proved the captured workbook, selected `System_Key`, exact
item/location/lot/Condition group, available quantity, zero-row staging table,
and the complete live inventory projection all stage successfully in a
disposable copy. The remaining live-only risk was re-entrant worksheet events
during the multi-cell staging write. Disposition staging now suppresses events
for the complete allocation/aggregate transaction and restores the caller's
prior event state; manual quantity synchronization likewise suppresses its own
write event and rebuilds the explicitly captured workbook rather than
`ActiveWorkbook`. The full-table matcher now exits on the first mismatched
attribute instead of evaluating unrelated row attributes. A same-handler
protected-write test recorded RED for the generic message and is GREEN with
stage, error number, source, and description context. Aligned Returns tests are
7/7, Slice 4o/4p/4q static contracts are 5/5, 6/6, and 6/6, packaged Returns is
5/5, packaged XLAM is 74/74, live workflows are 46/46, and the clean Release 1
chain is 30/30. Deterministic maintenance is 19/19 plus 11/11 with 150
components, 4,701 procedures, 965 scanner candidates, and an explicitly
reviewed Slice 4q procedure ceiling. The operator subsequently confirmed
RETURN, DUMP, reference concatenation, and Return Entries History posting
against `WHT7025AE`; Slice 4q visible disposition acceptance is complete.

### Slice 4r — ordinary receipt staging isolation and confirmation UI quieting

The next 2026-08-19 visible checkpoint selected Citric Acid for an ordinary
30-unit receipt at `CLEARVIEW`. **Add Selected** failed with the same generic
Excel error 1004 previously isolated on disposition staging. The same session
also exposed eighteen native Saving notifications after a successful batched
**Confirm Dispositions**, despite the processor remaining within its proven
three-persistence-save bound. This is a newly discovered Receiving acceptance
blocker inside Slice 4, not a change to the D14 identity or outbound
disposition contracts.

Required behavior:

- [x] ordinary `RECEIPT` staging suppresses Excel events for the complete
  Received Tally plus Aggregate Received transaction and restores the caller's
  prior event state on success or failure;
- [x] the real **Add Selected** form handler is protected by a focused action
  test that first proves detailed protected-write failure context and then
  proves successful receipt staging;
- [x] receipt failures identify the exact stage, error number, sanitized
  source, and description rather than returning generic error 1004 text;
- [x] the real **Confirm Writes/Confirm Dispositions** form handler enters the
  shared quiet-UI boundary before queue/processor/refresh work and restores the
  prior UI/event/calculation state afterward; and
- [x] batching remains authoritative: multi-row confirmation still persists
  once per safe artifact phase rather than once per row.

Gate:

- [x] same-handler RED is recorded before the implementation change;
- [x] focused Receiving tests prove ordinary receipt, Return/Dump, condition
  separation, exact identity, aggregate semantics, and bounded processor
  persistence;
- [x] packaged Receiving, XLAM, live role workflow, ordered Release 1 chain,
  and deterministic maintenance regressions remain GREEN; and
- [ ] the user visibly confirms an ordinary receipt stages successfully and
  that Confirm Writes/Confirm Dispositions no longer surfaces repeated Saving
  notifications against `WHT7025AE`.

Automated evidence recorded 2026-08-19: the real-form receipt action was RED
before implementation and GREEN afterward. The aligned Receiving range is
14/14; Slice 4o/4p/4q static contracts are 5/5, 8/8, and 6/6; packaged
Receiving/Returns is 5/5; packaged XLAM is 74/74; live role workflow is 46/46
with `QuietDuring=True` and `QuietRestored=True`; and the ordered Release 1
chain is 30/30. Deterministic maintenance is 19/19 and the Shipping/Boxing
regression is 11/11 with 150 components, 4,701 procedures, 8 literal
`Application.Run` targets, 47 unresolved expressions, and 184 duplicate-body
candidates. The 88-line bounded growth in the already oversized Receiving form
and controller is an explicit Slice 4r exception for the same-handler test,
stage diagnostics, event cleanup, and quiet-UI restoration; no component,
procedure, dynamic-call, or duplicate-body count increased.

### Slice 4s — Shipping exact-key reserve and active component projection

The 2026-08-20 visible Shipping checkpoint proved Box Maker could create
shippable boxes, then exposed three related Release 1 defects. Box Designer's
Component inventory included zero-balance durable entities that appeared as
unexpected duplicates. Shipping **Add** carried an exact `System_Key` but its
local reserve apply boundary still required `TOTAL INV/SHIPMENTS/ROW`. Required
Shipping and Boxing persistence also surfaced repeated native Saving notices.
These are newly discovered blockers within Slice 4; they do not change D14 or
the rule that distinct positive physical entities retain distinct keys.

Required behavior:

- [x] the real Shipping **Add** action reserves inventory by exact immutable
  `System_Key` on the current schema with no managed `ROW` header;
- [x] current-schema reserve validation preserves available quantity, existing
  locks, and inventory-floor enforcement before changing local staging;
- [x] Box Designer omits nonpositive component balances and removes only
  repeated projections of the same exact `System_Key`;
- [x] distinct positive entities remain separate component choices even when
  their displayed SKU, item, location, and description match;
- [x] Shipping Add, Box Designer save, and Box Maker Make/Unmake retain a
  shared quiet-UI boundary across required persistence; and
- [x] the quiet boundary hides Excel's status bar during the action and restores
  its previous setting with events, alerts, calculation, and screen updating.

Gate:

- [x] focused RED records the legacy `ROW` rejection, zero-balance picker rows,
  and missing Shipping quiet boundaries before implementation;
- [x] the same public `ShipmentsFormCommitLine` action passes on a `ROW`-free
  exact-key workbook and preserves the staged key and reservation quantity;
- [x] merge/exact-key/board regressions are 3/3, Slice 4s contracts are 6/6,
  Shipping/Boxing is 11/11, final control acceptance is 12/12, workflow
  readiness is 18/18, deployed live-role workflows are 46/46, and the ordered
  Release 1 full chain is 30/30;
- [x] deterministic maintenance is 19/19 with 150 components, 4,702 procedures,
  965 scanner candidates, 8 literal `Application.Run` targets, 47 unresolved
  expressions, and 184 duplicate-body groups; and
- [ ] the user visibly confirms the Box Designer list, Shipping Add, and native
  Saving-notification behavior against `WHT7025AE`.

The generic visible package inspector opened all five add-ins and passed 32/34;
its two failures are stale expectations for retired `AggregateBoxBOM_Log` and
`AggregatePackages_Log` support sheets, not active form/action failures. The
deployed live-role suite is the package-level gate for this slice and passed
46/46.

### Slice 4t — Shipping canonical post-send projection and bounded publication

The 2026-08-20 operator checkpoint completed a five-package shipment and proved
the canonical transaction correct: Inventory Viewer changed the selected box
from 100 to 95. The Shipping form nevertheless remained at NAS Inv 100,
Projected Inv 100, Locked 0. The action reported about 84 seconds total, about
58 seconds in `RunBatch`, and twenty native Saving notices. This is a newly
discovered Slice 4 projection/performance blocker, not an inventory-identity or
processor-application failure.

Root cause:

- `RunShippingRuntimeQueueRefresh` processed the queued event but returned
  without refreshing the captured operator workbook;
- `mBtnSend_Click` recalculated Projected Inv from its pre-send `mShippables`
  array instead of reloading the canonical read model;
- timer AutoSync could disarm after completed overlays were evicted, so it did
  not guarantee the missing refresh; and
- Shipping called `PublishInventorySnapshotBridge` after `RunBatch`, even
  though `RunBatch` had already generated/published the canonical snapshot.

Required behavior:

- [x] after processor success, refresh the captured Shipping operator workbook
  from the processor-generated canonical snapshot;
- [x] reload form shippables after that refresh and before deriving Projected
  Inv and Locked;
- [x] remove the redundant second snapshot publication while retaining the
  processor's three durability saves;
- [x] report nonzero read-model refresh timing separately from batch timing;
- [x] remove the duplicate legacy three-attempt shipment-stage cleanup call;
  and
- [x] visibly confirm the still-open Shipping form renders the same deducted
  NAS Inv, Projected Inv, and Locked values as Inventory Viewer. The follow-up
  six-package run completed in about 65 seconds, including about 49 seconds in
  the processor batch and 836 milliseconds in read-model refresh; a one-row
  Add still surfaced four Excel-native Saving notices.

Gate:

- [x] focused behavioral RED was 0/4: no canonical read-model call, no
  shippables reload, no public-form visible evidence, and duplicate legacy
  cleanup;
- [x] focused GREEN is 4/4;
- [x] the public form-action live workflow refreshes the same operator workbook
  from `starting quantity - shipped quantity` and is 47/47;
- [x] packaged XLAM validation is 74/74, Shipping/Boxing is 11/11, final control
  acceptance is 12/12, workflow readiness is 18/18, and the ordered Release 1
  full chain is 30/30; and
- [x] deterministic maintenance is 19/19 with 150 components, 4,702 procedures,
  964 scanner candidates, 8 literal `Application.Run` targets, 47 unresolved
  expressions, and 184 duplicate-body groups.

### Slice 4u — Shipping consolidated persistence feedback

The 2026-08-20 visible retest accepted the Slice 4t canonical form refresh, but
confirmed that Excel still displays native Saving progress windows around the
required NAS workbook writes. This is feedback noise, not failed persistence:
the one-row Add must durably save the warehouse inbox and reservation ledger,
and Shipments Sent must also retain the processor's three durability saves.
Excel-native progress windows cannot be moved into a VBA form. invSys therefore
provides one authoritative persistence summary in its existing message/status
surface and removes only avoidable per-row ledger saves.

Required behavior:

- [x] Add/Update/Remove and To Shipments append one `Persistence summary:` line
  after their required inbox and reservation-ledger writes;
- [x] Shipments Sent includes the same summary in the existing form status and
  final action dialog, including the processor's reported durability-save count;
- [x] a multi-row reservation action opens and saves the reservation ledger
  once rather than once per selected shipment row;
- [x] the processor's inventory, outbox, and inbox durability saves remain
  unchanged; and
- [x] the quiet-UI boundary remains active across each public form action, while
  documentation distinguishes invSys feedback from Excel-native progress UI.

Gate:

- [x] focused RED was 1/4: the processor durability contract existed, while Add
  summary, Shipments Sent summary, and batched reservation persistence were
  absent;
- [x] focused GREEN is 4/4 through the real Add and Shipments Sent callbacks;
- [x] packaged XLAM validation is 74/74, live role workflows are 47/47, and the
  ordered Release 1 full chain is 30/30;
- [x] Shipping/Boxing is 11/11, final control acceptance is 12/12, and workflow
  readiness is 18/18; and
- [x] reviewed cleanup is 11/11 with the procedure baseline unchanged at 4,702;
  static evidence remains 150 components, 964 scanner candidates, 8 literal
  `Application.Run` targets, 47 unresolved expressions, and 184 duplicate-body
  groups.

Visible follow-up:

- [ ] confirm the new `Persistence summary:` line appears after one Add and one
  Shipments Sent action, and record Excel-native Saving notices separately.

### Slice 4v — Receiving and Production consolidated persistence feedback

The 2026-08-20 follow-up requests the Shipping Slice 4u feedback behavior in
Receiving and Production as well. Excel-native Saving progress windows cannot
be reparented into VBA forms. The operator contract is therefore one
authoritative `Persistence summary:` in the existing form status surface while
required NAS durability remains intact.

Required behavior:

- [x] the real Receiving **Confirm Writes/Confirm Dispositions** callback
  appends one persistence summary after its batched inbox, processor, refresh,
  and cleanup work;
- [x] the real Production **Complete Run** callback appends one persistence
  summary after consume/complete event and processor work;
- [x] Production completion uses the shared quiet-UI boundary and restores the
  prior Excel UI state on success or error;
- [x] shared Receiving/Production run-and-refresh does not publish a second
  snapshot after the processor has generated the canonical snapshot; and
- [x] processor inventory, outbox, and inbox durability saves remain unchanged.

Gate:

- [x] focused behavioral RED was 0/4 before implementation;
- [x] focused GREEN is 4/4 through the real Receiving and Production callbacks;
- [x] Receiving aggregate/save batching is 8/8, disposition is 6/6, Receiving
  stabilization is 10/10, Production layout is 8/8, final control acceptance
  is 12/12, workflow readiness is 18/18, and Shipping/Boxing is 11/11;
- [x] packaged XLAM validation is 74/74, live role workflows are 47/47, and the
  ordered Release 1 full chain is 30/30; and
- [x] reviewed cleanup is 11/11 with 150 components, 4,702 procedures, 964
  scanner candidates, 8 literal `Application.Run` targets, 47 unresolved
  expressions, and 184 duplicate-body groups.

Visible follow-up:

- [ ] confirm one Receiving receipt, one Receiving disposition, and one
  Production Complete Run each produce one form-status persistence summary;
  record any remaining Excel-native Saving notices separately.

### Slice 4w — operator responsiveness and read-only event visibility

The 2026-08-20 follow-up reported a remote NAS connection that eventually
succeeded after Excel appeared frozen, truncated multi-reference aggregates,
continued Office-native Saving windows, and the need for a Viewer Events page.

Required behavior:

- [x] manual and ribbon Server Sign In render progress and yield once before
  the synchronous Windows SMB authentication call, then restore Excel UI state;
- [x] Receiving keeps its fixed-height aggregate rows and exposes the complete
  selected concatenated reference in a locked multiline detail surface that
  clears with staging;
- [x] Inventory Viewer exposes exactly Inventory and Events through the same
  reusable modeless form, with no designer/default `Tab1` or `Tab2` pages;
- [x] canonical inventory events are copied into the published snapshot as
  read-only `tblInventoryEvents` data, preserving Core/Domain authority;
- [x] operator-facing Events rows correspond to meaningful control actions and
  exclude internal processor/staging mechanics that occur while completing an
  action;
- [x] Events covers Receipt, Return, Dump, Box Made/Unboxed, Shipped, and Shipping
  **Remove** (`SHIP_RELEASE`, returning locked inventory to warehouse
  availability), supplemented by current saved Box Designs and actual current
  Held Shipments. Internal `SHIP_RESERVE` rows written by ordinary Shipping Add
  are not rendered as **Shipment Held**; Production labels are deferred to the
  Production review; and
- [x] Receiving/Returns, Production, and Shipping paint a form-owned pending
  status before persistence. Native Office Saving windows remain separate and
  required durability saves are not removed;
- [x] Viewer anchoring does not enforce form dimensions while the native window
  is minimized or maximized, and residual error 384 is contained so restore can
  reapply the layout;
- [x] Events renders readable timestamps and explicit Refresh replaces the
  already-open list with the newest published event projection; and
- [x] Events defaults to All dates on first use, explicit Refresh combines
  Search with a rolling Day, Week, Month, or typed positive whole-number-of-days
  window, and the last valid applied range is restored per Windows user when a
  new Viewer form or later Excel session opens, while Inventory remains
  unaffected; and
- [x] the normative specification records a later comprehensive Event Viewer as
  post-R1 design work rather than expanding the bounded R1 projection in place.

Gate:

- [x] focused behavioral RED was 0/5 before the original implementation;
  the Shipping Remove follow-up was 5/6, and the visible duplicate-tab/routing
  follow-up was 6/7 with only `Viewer.Tabs.ExactlyInventoryAndEvents` failing;
- [x] focused GREEN is 7/7;
- [x] the resize/readability/refresh follow-up RED was 7/9, with only native
  window-state guarding and readable timestamp refresh failing; GREEN is 9/9;
- [x] the rolling-date-filter follow-up RED was 9/10, with only
  `Viewer.Events.RollingDateFilters` failing; GREEN is 10/10;
- [x] packaged Viewer validation proved Day=2, Week=3, Month=4, custom
  14-day=3, and All=5 against the same five-row publication while leaving the
  snapshot hash unchanged;
- [x] the remembered-range follow-up RED was 10/11, with only
  `Viewer.Events.RemembersDateFilter` failing; GREEN is 11/11, and packaged
  validation restored custom 14 days after unloading and recreating the Viewer
  while restoring the machine's prior test preference afterward;
- [x] the internal-reservation-noise follow-up RED was 11/12, with only
  `Viewer.Events.ExcludesInternalReservation` failing; GREEN is 12/12, and the
  packaged public Viewer action hid the zero-delta `SHIP_RESERVE` fixture while
  retaining Shipping **Remove** and all range-filter behavior;
- [x] Receiving/Production feedback is 4/4, Shipping feedback is 4/4,
  Receiving aggregate/save batching is 8/8, disposition is 6/6, Receiving
  stabilization is 10/10, Shipping/Boxing is 11/11, and Production layout is
  8/8;
- [x] packaged XLAM validation is 74/74, live role workflows are 47/47, and the
  clean ordered Release 1 full chain is 30/30; and
- [x] reviewed cleanup is 11/11 with 150 components, 4,729 procedures, 965
  scanner candidates, 8 literal `Application.Run` targets, 47 unresolved
  expressions, and 185 duplicate-body groups. Procedure growth is the protected
  Events/progress implementation; the other ratchets remain improved.

Visible follow-up:

- [ ] confirm pre-connection progress is visible on the remote NAS path, read a
  four-reference aggregate in **Selected references**, and verify Viewer Events
  shows current receipt/disposition/boxing/shipping activity;
- [ ] count any Office-native Saving windows separately from the invSys pending
  and completion text. They cannot be moved into a UserForm.

### Slice 4x -- reusable Production Processes and Recipe graphs

This deliberate priority change temporarily supersedes the remaining Slice 4w
visual checkpoint order. It reconciles Production with Architecture v4.11 D15
before any conflicting VBA, schema, form, build, or runtime behavior changes.
After Slice 4x packaged evidence is GREEN, visible saved-workbook/NAS Production
testing resumes together with the still-open Slice 4 visual polish checks.

Required behavior:

- [x] replace the existing Recipe Builder top-level page with **Process
  Designer** and **Recipe Designer**, while retaining **Ingredients
  Assignment**, **Production Run - List**, and the experimental out-of-scope
  **Production Run - Tree**;
- [x] Process Designer creates, edits, validates, saves, releases, obsoletes,
  and reuses named immutable Process versions through the headless Designs
  Domain;
- [x] every Process declares input requirements, ordered instructions, one or
  more output definitions, positive quantity or percentage/yield bases, and
  valid UOMs; every output has its own item/design identity and every Process
  has at least one output;
- [x] Ingredients Assignment maps each Process requirement to acceptable
  managed item/SKU alternatives that are versioned with the Process;
- [x] Recipe Designer selects exact released Process versions, connects any
  output to compatible downstream requirements, supports multiple outgoing
  connections and unconnected finished/co-product balances, controls execution
  order, and rejects unresolved inputs, incompatible UOM/item/design identity,
  quantity over-allocation, missing/unreleased definitions, and cycles;
- [x] Production Run - List selects one released Recipe version, preserves
  inclusive scaling from `0.001%` through `1000%`, resolves acceptable
  alternatives against current inventory, and allocates exact available
  `System_Key` entities without aggregate identity substitution;
- [x] each Process output instance receives a distinct new `System_Key` before
  queueing; routed intermediate output is created under that key and consumed
  downstream by the same exact key, while unconsumed outputs remain managed
  finished/co-product inventory;
- [x] completion preserves inventory sufficiency checks, deterministic process
  order, correlated consume/complete events, canonical processor authority,
  persistence summaries, and read-only Viewer event labels for Production
  input consumption and output creation;
- [x] saved Production operator-workbook reuse, captured-workbook form binding,
  Operations launcher behavior, headless Core/Domain packages, and current
  Seed/Receiving/Shipping/Boxing/Viewer behavior remain unchanged; and
- [x] legacy single-builder recipes are never a silent runtime fallback when
  Designs are enabled. Any retained import is an explicit design-definition
  conversion into valid Process and Recipe versions.

D13 RED sequence:

1. [x] Add a focused source/packaged surface test entering through
   `mProduction.BtnOpenProductionForm`; record RED because the form still has
   one Recipe Builder page and lacks Process/Recipe designer handlers.
2. [x] Add Designs Domain tests for Process lifecycle, minimum one output,
   ingredient alternatives, Process reuse, Recipe graph resolution,
   compatibility, quantities/order, and circular dependencies; record RED for
   the missing event types/projections/validators.
3. [x] Add the packaged operator form-action path for Process
   save/release/obsolete, Recipe select/connect/order/save/release/obsolete, and
   ingredient assignment; service-only calls are supplemental and do not
   replace these handler tests.
4. [x] Extend the packaged two-batch run test through the actual recipe select,
   scale, palette Apply, Check In, Complete Run, Refresh, and Next Batch
   handlers. Prove `0.001%`, `100%`, and `1000%`, insufficiency, stale
   allocation, multi-output keys, intermediate routing, and co-product balance.
5. [x] Add processor/Inventory Domain RED for correlated multi-output
   `PROD_COMPLETE` envelopes and exact-key downstream `PROD_CONSUME`, followed
   by snapshot/operator/Event projection assertions.

Implementation order after meaningful RED:

1. [x] implement Designs Domain Process/Recipe lifecycle events, schemas,
   projections, validation, and read APIs;
2. [x] implement typed Production editor/controller contracts and replace the
   single Builder page without weakening captured-workbook binding;
3. [x] implement Recipe graph validation and ingredient-alternative persistence;
4. [x] implement typed multi-Process run plan/session and multi-output completion
   envelopes with one new key per output;
5. [x] implement processor/Inventory Domain application and published Production
   event visibility; and
6. [x] refactor only while focused and preserved regression ranges stay GREEN.

Slice 4x Designs Domain evidence: lifecycle/schema/graph RED and GREEN are
recorded in `tests/integration/plan022_slice4x_designs_domain_*_results.md` and
`tests/integration/plan022_slice4x_designs_lifecycle_graph_*_results.md`.
Released Process/Recipe read APIs and their Domain -> Core -> Operations
primitive bridge RED/GREEN are recorded in
`tests/integration/plan022_slice4x_designs_read_api_results.md`. The current
reusable range is 17/17 GREEN, packaged five-XLAM validation is 74/74 GREEN,
and static candidates are 967 versus the 968 baseline with literal
`Application.Run` 8/8, duplicate-body groups 185/185, and unresolved dynamic
calls improved from 47 to 45. The packaged public-launcher form-action RED and
GREEN are recorded in
`tests/integration/plan022_slice4x_form_action_results.md`: the GREEN proves
five pages, captured-workbook reuse, Process save/release/obsolete/reuse,
assignment-backed Process versioning, and Recipe connect/order/save/release/
obsolete through the actual click handlers. Five-page minimum/default/expanded
and native maximize/restore geometry is GREEN. The packaged reusable List-run
GREEN is recorded in
`tests/integration/plan022_slice4x_reusable_production_results.md`; it proves
`0.001%`, `100%`, and `1000%`, insufficiency and stale-key rejection, two
batches, six distinct output keys, exact intermediate consumption, retained
co-products including a 20%-of-10-unit yield basis, correlated processor
persistence, and public Viewer Production
event labels. Packaged XLAM/Ribbon validation is 74/74 and 142/142, deployed
live-role workflows are 47/47, and the ordered Release 1 chain is 30/30.
Deterministic maintenance is 19/19 with 152 components, 4,987 procedures,
1,033 candidates, 8 literal `Application.Run` targets, 45 unresolved dynamic
calls, and 189 duplicate-body groups; the deliberate Slice 4x growth has a
bounded reviewed exception at 13/13. A second clean Excel process reopens the
same saved station-local Production workbook and loads the persisted exact
released Recipe through the actual Run List **Load** handler; focused source is
4/4 and packaged restart runtime is 2/2. Visible dedicated-NAS acceptance
remains open. The current package also passes dedicated-NAS launcher readiness
16/16 across two clean sessions with zero canonical changes from launcher use.
That gate first reported 14/16 because Production list reads dirtied and saved
the Designs workbook; idempotent schema formatting and save-on-dirty-only
resolution corrected the regression under a 6/6 focused source contract.

Automated gate:

- [x] focused RED and GREEN reports distinguish pre-implementation behavior;
- [x] packaged Process/Recipe form actions and two consecutive List batches are
  GREEN through actual operator handlers;
- [x] Process/Recipe replay, idempotency, lifecycle, graph, scaling,
  insufficiency, exact-key allocation, multi-output, persistence, snapshot, and
  Viewer Event assertions are GREEN;
- [x] full five-XLAM compile/load/Ribbon validation and saved-workbook restart
  evidence are GREEN;
- [x] current launcher, Seed, Receiving, Shipping, Boxing, Viewer, full-chain,
  maintenance, dynamic-call, and bloat ratchets do not regress without a
  reviewed exception; and
- [x] Architecture v4.11, this plan, and `invSys-Controls-v1.md` match the
  implemented operator wording and acceptance state.

Visible Production acceptance after packaged GREEN:

- [ ] against the dedicated NAS test warehouse, create/release and reuse one
  named multi-output Process, assemble and release a multi-Process Recipe,
  prove a rejected unresolved/circular edit, assign acceptable alternatives,
  and complete two saved-workbook List runs at representative scales;
- [ ] confirm each produced output appears under a distinct new `System_Key`,
  routed intermediate inventory is consumed exactly, co-product inventory
  remains visible, persistence summary appears once, and Viewer Events shows
  the correlated Production actions; and
- [ ] maximize/restore the revised five-page form and confirm readable,
  non-overlapping Process Designer, Recipe Designer, Ingredients Assignment,
  Run List, and experimental Run Tree layouts.

### Slice 4y -- Process formulation worksheet round-trip and generated IDs

The 2026-08-23 visible checkpoint accepted **Upload Data Set**, **Seed Demo
Inventory**, and Inventory Viewer visibility of the uploaded sample. Viewer
therefore remains a preserved GREEN regression. The checkpoint found that
Process Designer exposes unexplained **Basis** fields and requires
operator-entered/GUID identities. The operator requested the existing
three-character Base-36 identity convention and a formula-driven Process editor
inside the captured saved Production workbook. This is a deliberate D15
contract refinement within Slice 4 saved-workbook/NAS acceptance; it does not
reopen the GREEN launcher, Viewer, or canonical inventory contracts.

Required behavior:

- [x] Process, Recipe, Requirement, and Output IDs are automatically allocated
  as locked three-character uppercase Base-36 values from `001` through `ZZZ`,
  with `000` reserved and collision checks in the applicable namespace;
- [x] replace operator-visible **Basis** wording with **Batch basis quantity**
  for percentage inputs and **Yield basis quantity** for percentage outputs;
  at 100% scale, quantity is `Percent / 100 * basis`, before run scaling;
- [x] add one Process Designer toggle action that sends the current draft to a
  uniquely named structured table in the exact captured
  `<WarehouseId>.Production.Operator.xlsm`, then retrieves that same table on
  the next action and removes it only after successful validation;
- [x] the worksheet table carries the Process header, requirements, outputs,
  and instructions; same-UOM input quantity rows use structured formulas to
  calculate the shared batch basis and percentages totaling 100%;
- [x] retrieval rejects unresolved identity, incomplete rows, incompatible or
  mixed UOM percentage formulations, invalid quantities, and missing outputs,
  leaving the table intact for correction;
- [x] editing a saved/released Process through the worksheet creates the next
  immutable DRAFT version; the worksheet never becomes Designs Domain or
  inventory authority; and
- [x] form close/reopen and saved workbook reopen rediscover the exact
  outstanding table without using `ActiveWorkbook`, and the operator can repeat
  send/retrieve after a successful round trip.

D13 RED sequence:

1. [x] Enter through `mProduction.BtnOpenProductionForm`; record RED because a
   new Process receives a GUID identity, line IDs are manual, ID controls are
   editable, and raw **Basis** wording remains.
2. [x] Invoke the actual Process Designer worksheet toggle handler; record RED
   because no uniquely bound formulation table, percentage formulas, or
   retrieve/delete path exists.
3. [x] Add focused save/reopen, mixed-UOM, invalid-table, and immutable-version
   tests through that same handler. Service-level formula/parser tests
   supplement but do not replace the packaged action proof.

Gate:

- [x] normative Architecture v4.11, this plan, and the controls catalog are
  updated before conflicting implementation;
- [x] focused Process handler RED/GREEN is recorded while Viewer regressions
  remain GREEN;
- [x] the sugar/flour/baking-powder/water example produces a 611.2 lb basis and
  displayed percentages 16.4%, 32.7%, 1.8%, and 49.1% (100.0% total);
- [x] successful retrieval updates the form draft and deletes only its owned
  temporary table; failed retrieval preserves the table and prior form draft;
- [x] packaged Process lifecycle/reuse, two-batch List run, saved-workbook
  restart, Viewer, XLAM/Ribbon, dedicated NAS, full-chain, static maintenance,
  dynamic-call, and bloat regressions remain GREEN; and
- [x] the operator visibly confirms uploaded inventory visibility; and
- [ ] the operator visibly confirms one create/edit/retrieve/re-edit Process
  worksheet round trip against the dedicated NAS test warehouse.

Automated evidence on 2026-08-23: focused source 6/6; packaged Process action
and clean restart 2/2; packaged XLAM 74/74; Ribbon/compile 142/142; live role
47/47; ordered Release 1 chain 30/30; launcher contracts 24/24; dedicated NAS
launcher safety 16/16 across two clean sessions; deterministic static baseline
19/19; and reviewed cleanup 13/13. The dedicated NAS run observed zero
canonical file changes from launcher use. Only the visible worksheet/layout UAT
remains open.

### Slice 4z -- multi-table Process import workbench

The 2026-08-24 visible checkpoint accepted creation of the first table on the
**invSys Process Editor** sheet and deliberately replaced the Slice 4y toggle
contract. Operators need to create and reformat many Process tables, including
CSV exports from other systems, before importing selected definitions into
invSys. This is a D15 contract change; Slice 4y remains historical RED/GREEN
evidence for the first working round trip.

Required behavior:

- [x] replace the toggle with separate **Create Process Table** and **Retrieve
  Selected Process** form actions;
- [x] allow multiple uniquely named Process tables in the captured saved
  Production workbook, with metadata owned by each table and no outstanding-
  table singleton restriction;
- [x] retrieval targets the table containing the selected captured-workbook
  cell, deletes only that table after success, and preserves every table on a
  failed import;
- [x] support paste/reformat of CSV rows into a generated table and add Record
  Type validation choices `INPUT`, `OUTPUT`, `INSTRUCTION`, and `ALTERNATIVE`;
- [x] make Percent and batch basis calculated columns for INPUT rows, including
  rows populated after table creation, and restore those formulas before
  validation;
- [x] remove operator-authored output Item Code, generate Design ID/version from
  Process/Output identity, and derive the inventory design SKU internally;
- [x] export/import existing Ingredient Assignment alternatives in each Process
  table without allocating inventory; and
- [x] invoke the existing Core item-search interaction when the operator selects
  an acceptable managed item cell, filling its managed item/SKU projection.

D13 RED sequence:

1. [x] Enter through `mProduction.BtnOpenProductionForm` and invoke the actual
   worksheet action: RED is the one toggle button and singleton-table rejection.
2. [x] Create two tables, select the first and retrieve it: RED is ambiguous
   global metadata and the prohibition on multiple outstanding tables.
3. [x] Populate a blank/pasted INPUT row after creation: RED is missing Percent
   formula; inspect Record Type for absent list validation and OUTPUT for manual
   Item Code/Design ID.
4. [x] Export an assigned requirement and select its acceptable-item cell: RED
   is absent assignment columns and no Core item-search path.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 are reconciled first;
- [x] focused source and packaged handler RED/GREEN are committed;
- [x] two or more tables survive save/reopen and selected retrieval removes only
  the successful table;
- [x] pasted rows receive formulas/validation/generated IDs and existing
  alternatives round-trip through the item-search-backed assignment columns;
- [x] Process/Recipe lifecycle, two-batch List run, Viewer, XLAM/Ribbon, live
  roles, full chain, dedicated NAS, static maintenance, dynamic-call, and bloat
  regressions remain GREEN; and
- [ ] the operator visibly accepts rapid multi-table creation, CSV reformat,
  selected retrieval, Ingredient Assignment search, and Recipe ordering.

Automated evidence on 2026-08-24: the focused source contract is 7/7 GREEN.
The packaged public-launcher gate is 2/2 GREEN: three simultaneous tables were
created, one selected table was retrieved while two remained, the remaining two
survived save/reopen, and clean-restart retrieval removed them one at a time.
The formulation evidence is 611.2 lb with calculated percentages
16.4/32.7/1.8/49.1 and a 100.0% total. Final regressions are packaged XLAM
74/74, Ribbon/compile 142/142, live roles 47/47, ordered Release 1 chain 30/30,
launcher contracts 24/24, dedicated NAS safety 16/16, deterministic static
baseline 19/19, and reviewed cleanup 13/13. The expanded-page layout RED found
one Retrieve-button/Description overlap; the polished rebuild is GREEN across
all five pages at minimum/default/expanded plus minimize/restore/maximize.
Operator-visible workbench acceptance remains required by the unchecked gate.

### Slice 4aa -- text-safe Process identities and bulk assignment import

The 2026-08-24 visible checkpoint accepted the managed Percent formulas and
exposed the next Process-workbench contract gaps. Excel displayed generated
`001` identities as `1`, Requirement ID remained operator-facing, UOM lacked
the Settings catalog dropdown, only one acceptable-item slot was available,
the Core search did not visibly open on cell entry, and Ctrl+click selections
could not retrieve several Process tables. This slice deliberately refines the
D15/Slice 4z worksheet contract while preserving its packaged evidence.

Required behavior:

- [x] preserve Process, Requirement, Output, and generated Design identity as
  text, including leading zeroes, and generate each INPUT Requirement ID;
- [x] apply the current Settings **Recipe UOM Catalog** as the UOM-column
  dropdown and reject values outside that catalog on retrieval;
- [x] replace the primary vertical alternative-entry layout with numbered
  **Acceptable Managed Item 1**, **2**, **3**, **4** managed item/SKU pairs on
  each INPUT row, with an action that appends another pair when needed;
- [x] open the existing Core item-search form when the operator enters any
  acceptable-item cell by mouse, Tab, or Enter, and commit into the matching
  numbered pair;
- [x] resolve every distinct Process table intersected by a captured-workbook
  Ctrl+click multi-area selection, validate it, save each successful definition
  as an immutable DRAFT through the public Process action, and delete only its
  confirmed table; and
- [x] preserve historical vertical `ALTERNATIVE` row import compatibility,
  captured-workbook binding, headless Core/Domain authority, and all prior GREEN
  regressions.

D13 RED sequence:

1. [x] Enter through `mProduction.BtnOpenProductionForm` and the actual create
   handler; RED is `1`/`D-1-1`, blank Requirement ID, and absent catalog UOM
   validation.
2. [x] Enter an acceptable-item cell through the worksheet selection event;
   RED is recognition without proof that the Core search form opened, plus the
   absence of numbered alternative pairs.
3. [x] Ctrl+click cells in two Process tables and invoke the actual Retrieve
   handler; RED is single-table resolution and no bulk DRAFT-save result.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 are reconciled before the
  changed implementation;
- [x] focused source and packaged public-handler RED/GREEN are recorded;
- [x] visible identities remain `001`-style text, Requirement IDs are automatic,
  UOM dropdown values match Settings, and numbered alternatives round-trip;
- [x] the actual selection event visibly opens Core item search for each
  numbered acceptable-item column;
- [x] Ctrl+click import saves and removes successful selected tables only, while
  failed/unselected tables remain; and
- [x] prior Production lifecycle/run, Viewer, packaged launcher, saved-workbook,
  NAS, XLAM/Ribbon, full-chain, static, dynamic-call, and bloat regressions
  remain GREEN before visible UAT resumes.

Automated evidence recorded 2026-08-24: focused source `8/8`, historical
Slice 4z source `7/7`, packaged public Production/restart `2/2`, packaged XLAM
`74/74`, Ribbon/compile `142/142`, live role workflows `47/47`, ordered Release
1 chain `30/30`, launcher contracts `24/24`, dedicated NAS `16/16`, static
determinism `19/19`, reviewed cleanup `13/13`, and five-page packaged layout
GREEN at minimum/default/expanded plus minimize/restore/maximize. The packaged
bulk result records `TextSafeIds=True`, `RequirementIds=True`,
`UomCatalog=True`, `NumberedAlternatives=True`, `AddedAlternative=True`,
`PickerOpened=True`, `MultiAreaSelection=True`, and `MultiTableDrafts=True`.
Visible operator acceptance of this refined worksheet workflow remains pending.

### Slice 4ab -- Process picker managed-inventory projection blocker

The 2026-08-24 visible checkpoint accepted Process-table creation, added tables,
and automatic item-search opening. The opened Production picker nevertheless
reported zero managed inventory rows while Inventory Viewer showed the seeded
managed inventory. This is a newly discovered Slice 4 acceptance blocker, not a
D15 architecture change: the Core picker still treated legacy `ROW` as its
required row identity even though Architecture v4.11 prohibits `ROW` and makes
`System_Key` the exact managed inventory identity.

Required behavior:

- [x] resolve Process acceptable-item search rows from the current captured
  warehouse inventory projection using exact nonblank `System_Key` identity;
- [x] show current managed inventory in the opened picker when the same
  inventory is visible in Inventory Viewer;
- [x] keep the Process worksheet assignment projection limited to acceptable
  managed item/SKU identity and never allocate a physical `System_Key` merely
  from picker selection; and
- [x] preserve captured-workbook binding, current-warehouse isolation, Core
  picker ownership, headless Domain authority, and all Slice 4aa GREEN behavior.

D13 RED sequence:

1. [x] Enter a numbered acceptable-item cell through the actual Production
   worksheet selection event after managed inventory is available.
2. [x] Record meaningful RED when the search form opens but its managed
   inventory result count is zero because the source requires `ROW`.
3. [x] Re-run the same public packaged action and require both
   `PickerOpened=True` and `PickerInventoryRows=True`.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 agree that `System_Key` is
  the only managed inventory row identity and no normative redesign is needed;
- [x] focused source and packaged public-handler RED/GREEN are recorded;
- [x] the packaged picker contains current managed inventory while preserving
  the exact captured workbook and warehouse context; and
- [x] Slice 4aa, packaged launcher, saved-workbook, NAS, full-chain, static,
  dynamic-call, and bloat regressions remain GREEN before visible UAT resumes.

Automated evidence recorded 2026-08-25: focused source `4/4`, historical Slice
4aa source `8/8`, historical Slice 4z source `7/7`, packaged public Production
and clean restart `2/2`, packaged XLAM `74/74`, Ribbon/compile `142/142`, live
role workflows `47/47`, ordered Release 1 chain `30/30`, launcher contracts
`24/24`, dedicated NAS `16/16`, deterministic static baseline `19/19`, and
reviewed cleanup `13/13`. The packaged result records both
`PickerOpened=True` and `PickerInventoryRows=True`. On 2026-08-25 the user
visibly confirmed that current managed inventory appears in the search tool;
selection into the numbered INPUT item/SKU pair remains part of final workflow
acceptance.

### Slice 4ac -- Process OUTPUT managed-item picker

The 2026-08-25 visible checkpoint accepted that Production Item Search now
shows current managed inventory. It then found that the same search interaction
is absent for OUTPUT rows. The operator clarified that every Process OUTPUT is
a managed inventory item. This is a deliberate D15 refinement: INPUTs continue
to own numbered Ingredient Assignment alternatives, while each OUTPUT must own
one picker-selected managed item/SKU in addition to its generated Design
identity.

Required behavior:

- [x] entering an OUTPUT **Name** cell by mouse, Tab, or Enter opens the same
  Core Production Item Search used by INPUT acceptable-item cells;
- [x] picker commit fills the visible output Name and a hidden system-managed
  **Output SKU** value; the operator never types or maintains an Item Code;
- [x] retrieve/save/release requires and persists that selected output SKU,
  while Output ID and Design ID/version remain generated;
- [x] selecting an output catalog item never captures or reuses an inventory
  `System_Key`; run completion creates every output entity under its own new
  exact key; and
- [x] preserve captured-workbook binding, current-warehouse isolation, Core
  picker ownership, headless Domain authority, numbered INPUT alternatives,
  and all Slice 4ab GREEN behavior.

D13 RED sequence:

1. [x] Create a Process table through the public form action and enter an
   OUTPUT Name cell through the actual worksheet selection event.
2. [x] Record meaningful RED because the cell is not a picker target and the
   Core commit handler only resolves numbered INPUT alternative columns.
3. [x] Commit a current managed item through the same picker handler, retrieve
   the table through the public form action, and require the output Name/SKU to
   round-trip without any physical `System_Key`.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 agree on the distinct
  managed output SKU, generated Design identity, and new execution-time key;
- [x] focused source and packaged public-handler RED/GREEN are recorded;
- [x] Slice 4ab plus applicable packaged, launcher, NAS, full-chain, static,
  dynamic-call, and bloat regressions remain GREEN; and
- [ ] visible operator confirmation proves OUTPUT search and retrieval before
  Production workflow acceptance resumes.

Automated evidence recorded 2026-08-25: focused RED `1/6`, focused GREEN `6/6`,
historical Slice 4ab `4/4`, Slice 4aa `8/8`, and Slice 4z `7/7`; packaged public
Production callback and clean restart `2/2` with `OutputPickerOpened=True`,
`OutputPickerCommitted=True`, `OutputSkuHidden=True`,
`OutputSkuRoundTrip=True`, and `NoPhysicalKey=True`; packaged XLAM `74/74`,
Ribbon/compile `142/142`, live role workflows `47/47`, clean ordered Release 1
chain `30/30`, launcher contracts `24/24`, dedicated NAS `16/16`, deterministic
static baseline `19/19`, and reviewed cleanup/growth `13/13`.

### Slice 4ad -- Process picker INPUT/OUTPUT record-type reachability

The 2026-08-25 visible Slice 4ac checkpoint found that Production Item Search
opens from an INPUT row's **Acceptable Managed Item** cell but does not open
from the corresponding visible cell when Record Type is OUTPUT. Slice 4ac's
packaged proof targeted OUTPUT Name and therefore did not protect the operator's
same-column INPUT/OUTPUT workflow. This is a newly discovered acceptance
blocker and test-gap correction within D15.

Required behavior:

- [x] entering **Acceptable Managed Item 1** by mouse, Tab, or Enter opens the
  same Core Production Item Search for both INPUT and OUTPUT record types;
- [x] for INPUT, commit continues to fill the exact numbered managed item/SKU
  alternative pair;
- [x] for OUTPUT, commit fills the visible managed-item selector and hidden
  system-managed **Output SKU**, retaining a separately authored descriptive
  Output Name unless it is blank;
- [x] retrieve/save/release persists the OUTPUT SKU without treating it as an
  INPUT alternative or storing a source inventory `System_Key`; and
- [x] preserve OUTPUT Name convenience selection, current managed-inventory
  projection, captured-workbook binding, headless Domain authority, and all
  Slice 4ac regressions.

D13 RED sequence:

1. [x] Create a Process table through the public form action and populate an
   OUTPUT row through the existing formulation handler.
2. [x] Enter that OUTPUT row's **Acceptable Managed Item 1** cell through the
   actual worksheet selection event; record RED because the predicate accepts
   that column only for INPUT/REQUIREMENT rows.
3. [x] Commit a managed item through the real Core picker handler and retrieve
   through the public form action; require visible item/hidden Output SKU
   round-trip and no physical key.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 agree on the visible
  same-column INPUT/OUTPUT interaction;
- [x] focused and packaged public-handler RED/GREEN correct the Slice 4ac
  false-positive test target;
- [x] Slice 4ac and applicable packaged, launcher, NAS, full-chain, static,
  dynamic-call, and bloat regressions remain GREEN; and
- [ ] visible operator confirmation proves the picker opens for both INPUT and
  OUTPUT record types.

Automated evidence recorded 2026-08-25: focused source RED `5/6`; packaged
public-handler RED `0/2` with `OutputPickerOpened=False`,
`OutputPickerCommitted=False`, and `OutputSkuRoundTrip=False`; focused GREEN
`6/6`; packaged public Production callback and clean restart `2/2` with
`OutputPickerOpened=True`, `OutputPickerCommitted=True`,
`OutputSkuRoundTrip=True`, `OutputNameRetained=True`, and
`NoPhysicalKey=True`. Historical Slice 4ac/4ab/4aa/4z source contracts remain
`6/6`, `4/4`, `8/8`, and `7/7`. Packaged XLAM is `74/74`, Ribbon/compile
`142/142`, live role workflows `47/47`, clean ordered Release 1 chain `30/30`,
launcher contracts `24/24`, dedicated NAS `16/16`, deterministic static
baseline `19/19`, and reviewed cleanup/growth `13/13`.

### Slice 4ae -- Process picker cell boundary and table-wide row identity

The 2026-08-25 visible Slice 4ad retest confirmed OUTPUT picker reachability but
exposed two contract defects. Entering an OUTPUT **Name** cell still opened
Production Item Search, although managed-item selection belongs only in
**Acceptable Managed Item n** cells. Separately, worksheet automation allocated
INPUT and OUTPUT IDs from different namespaces, allowing the same three-digit
Base-36 ID to appear twice in one Process table. This is a deliberate Slice 4
acceptance correction under D15 and supersedes Slice 4ad's OUTPUT Name
convenience allowance.

Required behavior:

- [x] Production Item Search opens only from an **Acceptable Managed Item n**
  cell valid for that row type: any numbered pair for INPUT/REQUIREMENT, and
  pair 1 for OUTPUT;
- [x] OUTPUT Name and all non-assignment cells never open item search;
- [x] INPUT, REQUIREMENT, OUTPUT, and INSTRUCTION rows receive text-safe
  three-character Base-36 IDs from one table-wide namespace, with `000`
  reserved;
- [x] an existing valid unique row ID remains stable when other row types are
  entered later, while any pre-existing duplicate is reassigned automatically;
  and
- [x] preserve exact managed-item/SKU commit, Output SKU round-trip,
  captured-workbook binding, headless Domain authority, and prior GREEN
  regressions.

D13 RED sequence:

1. [x] Add a focused source contract requiring the cell boundary, shared ID
   namespace, and packaged public-handler evidence; record `0/6` RED.
2. [x] Through the public Process worksheet selection handler, prove OUTPUT
   Name is suppressed and OUTPUT Acceptable Managed Item 1 still opens.
3. [x] Through `HandleProductionChange`, assign OUTPUT first, then INPUT and
   INSTRUCTION rows; require uniqueness after each change and retention of the
   first assigned ID.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 define the corrected cell
  and row-identity boundaries before implementation;
- [x] focused and packaged callback/form-handler RED/GREEN are recorded;
- [x] applicable Production, launcher, NAS, full-chain, static, dynamic-call,
  and bloat regressions remain GREEN; and
- [ ] visible operator confirmation proves only Acceptable Managed Item cells
  open search and row IDs remain unique regardless of entry order.

Automated evidence recorded 2026-08-25: focused source RED `0/6`; packaged
ProductionReusable RED `1/2` with `OutputNamePickerSuppressed=False` and
`UniqueRowIds=False` while picker commit and restart remained functional;
focused GREEN `6/6`; packaged ProductionReusable GREEN `2/2` with
`OutputNamePickerSuppressed=True`, `UniqueRowIds=True`, and
`FirstAssignedIdRetained=True`. Historical Slice 4ad/4ac/4ab/4aa/4z/4y source
contracts remain `6/6`, `6/6`, `4/4`, `8/8`, `7/7`, and `6/6`. Packaged XLAM
is `74/74`, Ribbon/compile `142/142`, live roles `47/47`, ordered Release 1
chain `30/30`, launcher contracts `24/24`, dedicated NAS `16/16`, deterministic
static baseline `19/19`, and reviewed cleanup/growth `13/13`.

### Slice 4af -- Recipe identity initialization and editable version

The 2026-08-26 visible Recipe Designer checkpoint found that entering a Recipe
name and adding a released Process still left Recipe ID and Version blank.
**Save Draft** and **Release** consequently stopped at `Recipe ID, version, and
name are required.` The operator confirmed that Recipe ID is invSys-owned and
must be automatic, while Version must receive an automatic proposal but remain
editable. This is a deliberate D15 contract correction within Slice 4 saved-
workbook/NAS Production acceptance and supersedes the earlier locked Recipe
version control requirement.

Required behavior:

- [x] after the Production form initializes its saved Recipe list, a blank
  Recipe draft displays the next collision-checked three-character Base-36 ID
  and proposed version `1` without requiring **New Recipe**;
- [x] **New Recipe** and **Clear** generate the same blank-draft identity, and
  **Save Draft** and **Release** defensively restore either value if it is blank
  before validating the operator's Recipe graph;
- [x] Recipe ID remains locked and operator-inaccessible; Recipe Version is an
  editable positive whole-number field whose nonblank valid operator value is
  preserved through the actual Save Draft and Release handlers;
- [x] a saved Recipe ID/version remains immutable, and the Designs Domain
  rejects an unavailable or duplicate version rather than overwriting it; and
- [x] Process identities, Recipe graph validation, captured-workbook binding,
  packaged launcher reuse, headless Domain authority, and all prior GREEN
  Production and Release 1 regressions remain unchanged.

D13 RED sequence:

1. [x] Add the focused Slice 4af source contract and record `0/6` RED against
   the public Recipe Designer control and handler contract.
2. [x] Extend packaged Production evidence through the actual New Recipe, Save
   Draft, and Release click handlers to prove automatic ID/version, locked ID,
   editable Version, and retention of an operator-edited version.
3. [x] Rebuild and re-run the preserved Production, launcher, NAS, full-chain,
   static, dynamic-call, and bloat regression gates before visible UAT resumes.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 define the revised Recipe
  identity contract before implementation;
- [x] focused RED/GREEN and packaged public-handler GREEN are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation proves ID/version are already populated,
  Version can be edited, and Save Draft/Release no longer fail for missing
  generated identity.

Automated evidence recorded 2026-08-26: focused source RED `0/6` and GREEN
`6/6`; packaged Production public callback plus clean restart `2/2`. The actual
form paths record `RecipeIdentityInitialized=True`, `RecipeIdGenerated=True`,
`RecipeVersionGenerated=True`, `RecipeIdLocked=True`,
`RecipeVersionEditable=True`, and `EditedRecipeVersionRetained=True` after the
operator-path test changes Version from `1` to `9` and invokes **Save Draft**
and **Release**. Historical Slice 4y/4z/4aa/4ab/4ac/4ad/4ae source contracts
remain GREEN. Packaged XLAM is `74/74`, Ribbon/compile `142/142`, live roles
`47/47`, ordered Release 1 chain `30/30`, launcher contracts `24/24`, dedicated
NAS `16/16`, deterministic static baseline `19/19`, and reviewed cleanup/growth
`13/13`. Static metrics are 153 components, 5,081 procedures, and 1,038
candidates. Rebuilt packages are registered; visible Recipe Designer acceptance
remains open.

### Slice 4ag -- reusable-run actual output and list readability

The 2026-08-26 visible Production Run - List checkpoint completed the first
operator-created tea batch and exposed two related acceptance blockers. The
entered **Real Output** value `430` remained absent from the **Last** column,
and source inspection proved the reusable-run completion path ignored that
entry and created inventory at the planned `632`. The same view clipped exact
identity under narrow **System Key** and **Inventory ID** headers. This is a
deliberate D15 correction: reusable Production now makes per-output operator
actuals authoritative for created quantity while preserving the scaled design
quantity as the visible plan.

Required behavior:

- [x] **Actual Output** accepts and retains one positive operator quantity for
  each reusable output row through the actual output-row/textbox interaction;
- [x] **Complete Run** rejects any missing/nonpositive actual, and rejects a
  routed output whose actual is smaller than its committed downstream amount;
- [x] each `PROD_COMPLETE` creates the output's new exact `System_Key` entity at
  the operator-entered actual quantity while recording both planned/scaled and
  actual quantities in correlated event attributes;
- [x] after completion, **Last Actual** shows the entered quantity for the most
  recent batch, **Batch** shows its batch number, and **Planned** continues to
  show the scaled definition quantity; Next Batch retains Last Actual and
  clears only the new batch's staged actual entries;
- [x] palette, Inventory Check, and Production Output identity headers read
  exactly **System_Key** and have sufficient width for operator readability;
  Production Output no longer calls that identity **Inventory ID**; and
- [x] preserve exact-key input consumption, fresh output keys, intermediate
  routing, co-product balances, scaling, captured-workbook binding, packaged
  launcher reuse, headless Domain authority, and every prior GREEN regression.

D13 RED sequence:

1. [x] Add focused source/public-handler coverage and record `0/6` RED for the
   ignored actual quantity, blank Last value, and unreadable identity headers.
2. [x] Through the packaged public Production callback, select every output
   row, enter actual quantities through `txtOutputReal`, invoke Check In and
   Complete Run, and prove Last Actual plus exact created inventory quantity.
3. [x] Rebuild and re-run the preserved Production, launcher, NAS, full-chain,
   static, dynamic-call, and bloat gates before visible UAT resumes.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 define the corrected
  planned-versus-actual and `System_Key` display contract before implementation;
- [x] focused RED/GREEN and packaged public-handler GREEN are recorded;
- [x] applicable regressions remain GREEN; and
- [x] visible operator confirmation proves the completed tea batch's actual
  quantity appears under Last Actual and in managed inventory.

Automated GREEN recorded 2026-08-26: focused `6/6`, packaged Production public
callback plus clean restart `2/2`, packaged XLAM `74/74`, Ribbon/compile
`142/142`, live role workflows `47/47`, ordered Release 1 chain `30/30`,
launcher contracts `24/24`, dedicated NAS `16/16`, deterministic static
baseline `19/19`, and reviewed cleanup/growth `13/13`. Static metrics are 153
components, 5,092 procedures, and 1,038 candidates. The earlier operator batch
was already persisted at planned `632` by the superseded build; Slice 4ag does
not silently rewrite that inventory entity. Visible acceptance therefore uses
a newly completed batch or an explicitly authorized inventory correction.

Visible GREEN recorded 2026-08-26: the operator completed a new tea batch,
confirmed **Last Actual** on Production Run - List, confirmed the new managed
inventory entity, and accepted the reusable batch system behavior.

### Slice 4ah -- Admin inventory edit selection binding

The 2026-08-26 visible Admin **Edit inventory item** checkpoint selected
Filtered Water from the Inventory item combo, but Save reported **Choose an
inventory item to edit**. Source inspection proved the combo Change handler
treated a real dropdown selection as typed search text, rebuilt the list,
forced `ListIndex = -1`, and failed to bind the selected SKU even though the
selected display text remained visible. This is a defect against the existing
Admin control contract, not an Architecture v4.11 change.

Required behavior:

- [x] choosing a combo dropdown row immediately binds its catalog SKU and loads
  the same item's name, UOM, location, description, category, vendor, external
  code, quantity mode, and custom metadata;
- [x] typed search continues to filter without inventing an item selection;
- [x] choosing a search-results row and choosing a combo dropdown row converge
  on the same SKU-backed loader;
- [x] changing the selected item to **Utility** produces `TRACK_QTY=FALSE` and
  `ITEM_KIND=UTILITY` for that same item without requiring a numeric quantity;
  and
- [x] preserve Add Item, edit-reason audit, exact catalog SKU authority,
  packaged Admin startup, inventory projection, and every prior GREEN
  regression.

D13 RED sequence:

1. [x] Record focused `0/5` RED for the selection-clearing Change handler and
   missing packaged assertion.
2. [x] Exercise the real combo Change handler with a selected catalog row and
   prove SKU binding, field loading, and Utility readiness.
3. [x] Rebuild the five-package set and re-run Admin/package, Release 1, static,
   dynamic-call, and growth gates before visible UAT resumes.

Gate:

- [x] Existing Architecture v4.11 and controls establish catalog identity and
  non-counted Utility behavior; no normative architecture change is required;
- [x] focused RED/GREEN and packaged handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [x] visible operator confirmation saves Filtered Water as Utility without a
  false missing-selection message.

Automated GREEN recorded 2026-08-26: focused `5/5`, packaged XLAM/Admin handler
`75/75`, Ribbon/compile `142/142`, live role workflows `47/47`, ordered Release
1 chain `30/30`, launcher contracts `24/24`, dedicated NAS `16/16`,
deterministic static baseline `19/19`, and reviewed cleanup/growth `13/13`.
The packaged handler records `ComboSelected=True`, `FieldsLoaded=True`,
`UtilityReady=True`, and `ValidationReady=True`. Static metrics are 153
components, 5,096 procedures, and 1,038 candidates.

Visible GREEN recorded 2026-08-26: Filtered Water was selected, edited, and
saved as **Utility** without the false missing-selection message. The subsequent
Production view exposed a separate Utility projection defect handled by Slice
4aj; it does not reopen the accepted Admin selection/save contract.

### Slice 4ai -- Admin bulk inventory worksheet staging

The operator accepted Slice 4ah and deliberately expanded the Admin contract:
the ribbon launcher must describe both actions, and the Add/Edit form must let
an administrator create a worksheet table for a pasted or hand-built inventory
list and upload that selected table through invSys. Architecture v4.11's D14
Admin inventory worksheet workbench is authoritative for this change.

Required behavior:

- [x] rename the Admin ribbon button from **Add Inventory Item** to **Add/Edit
  Inventory Items** without changing its `ADMIN_MAINT` capability or public
  callback;
- [x] add visible **Create Inventory Table** and **Upload Selected Inventory
  Table** actions to `frmAddInventoryItem`;
- [x] bind both actions to the workbook captured when the public Admin launcher
  opened the form; never switch authority because another workbook activates;
- [x] create uniquely named tables on **invSys Inventory Editor**, permit
  multiple simultaneous tables and pasted/reformatted CSV rows, and preserve
  unknown columns as custom catalog fields;
- [x] expose normalized managed headers for ADD/EDIT, generated or exact Item
  Code, item fields, configured UOM, quantity mode/quantity, edit reason, and
  upload status/result; prohibit `ROW` and `System_Key`;
- [x] validate every pending row before the first authoritative write; generate
  text-safe ADD Item Codes, require exact existing Item Code plus reason for
  EDIT, and preserve counted/Utility/Service/Not-counted behavior;
- [x] apply through the existing Admin catalog/event/processor boundaries,
  record per-row outcome, skip already-successful rows, and leave failed or
  unprocessed rows available for correction; and
- [x] preserve the accepted single-item Add/Edit workflow, packaged launcher
  reuse, all five package roles, and every prior GREEN regression.

D13 RED sequence:

1. [x] Record focused RED for old ribbon wording, absent form controls/handlers,
   absent worksheet controller, and absent packaged action proof.
2. [x] Through the real create/upload form handlers, create a table in a
   captured scratch workbook and prove preflight of counted ADD, Utility ADD,
   and exact EDIT rows without substituting a direct service-only test.
3. [x] Rebuild the five-package set and re-run Admin/package, Ribbon/compile,
   Release 1, launcher, NAS, static, dynamic-call, and growth gates.

Gate:

- [x] Architecture v4.11, Plan 022, and the controls catalog define the changed
  contract before implementation;
- [x] focused RED/GREEN and packaged handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation creates, edits, and uploads an inventory
  table from the saved workbook.

Automated GREEN on 2026-08-26: the focused contract is `8/8`; packaged XLAM
actions are `76/76`, including the real form handlers with table creation,
whole-table preflight, counted/Utility/exact-EDIT cases, generated codes, and
row statuses; Ribbon/compile is `142/142`; live roles are `47/47`; the Release 1
full chain is `30/30`; source/default packaged launcher checks are `24/24` and
`3/3`; dedicated NAS acceptance is `16/16`; deterministic static checks are
`19/19`; and reviewed cleanup/growth ratchets are `13/13`. Static inventory is
154 components, 5,140 procedures, and 1,040 candidates. Visible acceptance
remains open.

### Slice 4aj -- Production batch history and Utility projection

The accepted reusable tea run proved actual-output persistence and then exposed
two related operator-display gaps. Production Output overwrites its one visible
row when Next Batch advances, and the exact-entity projection continues to show
a legacy numeric balance for a catalog item that Admin successfully changed to
Utility. This slice deliberately changes the D15 Production Run - List display
contract without weakening exact-key event identity or Inventory Domain
authority.

Required behavior:

- [x] retain one Production Output display row for each completed batch and
  Process output, preserving that batch's Last Actual, recall, and `System_Key`;
- [x] keep the active unfinished batch as a separate selectable row so Actual
  Output entry continues through the operator's existing list/text handlers;
- [x] replace the visible **Planned** heading with **Used Goods**, calculated as
  the scaled compatible input quantity consumed by that Process for the batch;
- [x] add **Process Total**, the cumulative actual quantity for the same
  Process/output/UOM across retained completed rows, without adding unlike UOMs;
- [x] append catalog `TRACK_QTY`, `ITEM_KIND`, and Category metadata to the
  read-only exact-entity query envelope and render **Utility** in Run palette
  Inv and Inventory Check for a Utility-managed item instead of its historical
  numeric balance;
- [x] preserve measured Utility usage in the required/used quantity columns,
  existing non-counted Inventory Domain apply behavior, exact-key allocation,
  sufficiency/stale checks for counted inventory, multi-output completion,
  batch scaling, saved-workbook binding, and prior GREEN regressions.

D13 RED sequence:

1. [x] Record focused behavioral RED for the eight-column overwritten output
   list, Planned heading, absent Process Total, and missing exact-entity
   quantity-mode metadata.
2. [x] Through the real packaged Production output-row/textbox, Check In,
   Complete Run, and Next Batch handlers, complete two batches and prove two
   retained completed rows plus the active batch, batch-specific keys/actuals,
   Used Goods, and cumulative Process Total.
3. [x] Rebuild the affected package set and re-run Production, Inventory Domain,
   launcher, NAS, Release 1, static, dynamic-call, and growth gates.

Gate:

- [x] Architecture v4.11, this plan, and controls v1 define the changed display
  and Utility projection contracts before implementation;
- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation shows multiple batch rows, Process Total,
  and **Utility** instead of `19400` for Filtered Water.

Automated GREEN recorded 2026-08-27: focused RED was `1/7` and focused GREEN is
`7/7`; the packaged Production callback plus clean restart is `2/2` and records
`BatchHistoryRows=True`, `ProcessTotal=True`, and `UtilityDisplay=True` through
the real output selection/textbox, Check In, Complete Run, Refresh, and Next
Batch handlers. Packaged XLAM and Ribbon validation are `76/76` and `142/142`,
live role workflows are `47/47`, the ordered Release 1 chain is `30/30`, the
dedicated two-session NAS gate is `16/16`, deterministic static evidence is
`19/19`, and reviewed growth/cleanup is `13/13`. Visible acceptance remains
open.

### Slice 4ak -- Admin-created inventory visibility and catalog dropdowns

The operator added Honey through **Add/Edit Inventory Items**. Honey appeared
in the form's catalog-backed Edit search but not in Inventory Viewer or the
Production Process item picker. Inspection shows the Add path queues
`MIGRATION_SEED` from `CreatePayloadItem`, leaving `System_Key` blank. The
catalog therefore contains Honey, but no managed inventory entity exists for
the entity-backed Viewer snapshot or Production exact-key picker. This is a
newly discovered Release 1 blocker and an existing D14 identity violation, not
a request to restore legacy migration behavior.

Required behavior:

- [x] the single-item Add handler and worksheet `ADD` use
  `CreateInventoryEntityPayloadItem` plus `QueueInventoryCreateEvent` with one
  fresh Admin-generated `System_Key` per row;
- [x] a counted item with positive starting quantity becomes an active managed
  entity and appears after Refresh in Inventory Viewer and the Production
  managed-item picker;
- [x] a catalog-only item made by the superseded Add path can be completed by an
  explicit Edit/Save with a positive target only when no entity exists, creating
  its first new key without mapping `ROW` or importing legacy inventory;
- [x] Default location and Category are editable dropdowns populated with
  distinct current catalog values, with the configured warehouse default also
  present in Default location;
- [x] preserve Utility/Service/Not Counted metadata, Admin worksheet staging,
  saved-workbook binding, immutable exact keys, snapshot publication, and all
  prior GREEN regressions.

D13 RED sequence:

1. [x] Record focused RED for the migration event, blank-key payload, text-box
   controls, and absent packaged Add/dropdown evidence.
2. [x] Exercise the real form submit handler and Admin creation service, then
   prove one exact managed entity is eligible in Viewer and Production query
   projections after processor/snapshot publication.
3. [x] Rebuild the five-package set and rerun focused, packaged Admin,
   Inventory Domain, Viewer/Production, Release 1, NAS, static, and growth gates.

Gate:

- [x] Architecture v4.11, Plan 022, and controls v1 define the corrected
  identity and dropdown contracts before implementation;
- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation shows Honey in Viewer and the Production
  picker and shows dropdown arrows/options for Category and Default location.

### Slice 4al -- Add/Edit managed-inventory deletion

The operator requested deletion directly from **Add/Edit Inventory Items**.
This is a deliberate Release 1 contract addition. It does not authorize
physical deletion of immutable entity identity or event history. The Admin
surface therefore provides an audit-preserving SKU retirement action that
removes every active exact entity for the selected item from active managed
inventory, including zero-quantity Utility/Service entities.

Required behavior:

- [x] expose **Delete Item** only in Edit mode after the catalog-backed search
  has bound an exact SKU;
- [x] require a second confirmation and nonblank deletion reason through the
  real button handler;
- [x] resolve every active entity for the selected SKU and queue one
  `ADMIN_INVENTORY_ADJUST` payload with a separate exact `System_Key` line per
  entity, reducing counted quantities to zero and writing zero-delta retirement
  evidence for non-counted entities;
- [x] rebuild those entities as `InventoryState=RETIRED`, mark the retained
  catalog record `RETIRED`, and omit the item from ordinary Edit search,
  Inventory Viewer inventory levels, and Production managed-item pickers after
  Refresh; and
- [x] never erase or reuse a key, physically delete catalog/entity/log rows, or
  add a worksheet bulk-delete verb in this Release 1 slice.

D13 RED sequence:

1. [x] Exercise the real form Delete handler after selecting a catalog-backed
   item and record RED because no delete control/action exists.
2. [x] Apply a counted and a Utility retirement through the Admin service and
   record RED until exact-key negative/zero-delta retirement lines rebuild as
   `RETIRED` and disappear from active inventory queries.
3. [x] Rebuild the five-package set and rerun focused, packaged Admin, Inventory
   Domain, Viewer/Production, Release 1, NAS, static, and growth gates.

Focused source RED was `1/7`: documentation was present while the form action,
Admin service, Inventory Domain retirement, active-projection filtering, and
packaged callbacks were absent. After first deployment, visible launch exposed
a compile defect because the new delete variables were declared in the adjacent
worksheet handler instead of `Add_InventoryItem`. The strengthened focused test
recorded declaration-scope RED at `7/8`; moving both declarations into the real
public Ribbon handler restored focused GREEN at `8/8`. The rebuilt packaged
gate is `79/79`, including the real Delete button handler and isolated counted
plus Utility retirement through the Inventory Domain. Preserved evidence is
Ribbon/compile `142/142`, live role workflows `47/47`, ordered Release 1
`30/30`, dedicated NAS runtime `16/16`, deterministic static `19/19`, and
reviewed growth/cleanup `13/13`.

Gate:

- [x] Architecture v4.11, Plan 022, and controls v1 define the audit-preserving
  deletion contract before implementation;
- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation shows the selected item is absent from the
  managed inventory list, Viewer, Production picker, and ordinary Edit search
  after Refresh while Events retain its administrative retirement evidence.

### Slice 4am -- Zero starting quantity for managed inventory creation

The operator confirmed that a managed item definition may be established before
inventory arrives or is produced. This deliberately supersedes Slice 4ak's
positive-quantity creation/visibility constraint. It does not weaken the
nonnegative inventory floor or make zero quantity equivalent to retirement.

Required behavior:

- [x] the real single-item **Add Item** handler accepts required numeric Starting
  Qty `0` and rejects values below zero;
- [x] inventory worksheet counted `ADD` accepts explicit numeric Quantity `0`,
  continues to reject blank/non-numeric and negative quantities, and passes the
  zero value through the real upload/apply handler;
- [x] `INVENTORY_CREATE` accepts counted zero and creates one active immutable
  `System_Key` entity instead of writing catalog-only state or `MIGRATION_SEED`;
- [x] active zero-quantity managed items remain visible after Refresh in managed
  inventory, Inventory Viewer, Receiving choices, and the Production managed-
  item picker, while retired items remain omitted; and
- [x] later receipt or Production output creates its own new physical entity
  key and all negative creation paths remain rejected.

D13 RED sequence:

1. [x] Exercise the real Add button handler with counted zero and record RED
   until it accepts zero while rejecting negative input.
2. [x] Exercise worksheet preflight/upload and isolated Inventory Domain apply
   with counted zero and record RED until both preserve an active exact entity.
3. [x] Rebuild the five-package set and rerun focused, packaged Admin/Domain,
   Viewer/Production, Release 1, NAS, static, and growth gates.

Focused RED was `1/8`: only the reconciled documentation passed while the form,
Admin service, worksheet upload, Domain apply, managed picker, Viewer, and
packaged evidence retained the positive-only rule. Focused GREEN is `8/8`.
The first packaged Domain run then exposed a rebuild-state defect at `80/81`;
the strengthened exact-entity test recorded `7/8` RED before zero entities were
made active. Final packaged action evidence is `81/81`, including the real Add
handler, zero counted worksheet upload, isolated active exact-key creation, and
negative rejection. Exact Production Run allocation continues to list positive
available entities only; Process Designer's managed-item picker uses the active
catalog projection and can therefore select a zero-stock definition safely.
Final regression evidence is packaged action `81/81`, Ribbon/compile `142/142`,
live roles `47/47`, ordered Release 1 `30/30`, dedicated NAS runtime `16/16`,
deterministic static `19/19`, and reviewed growth/cleanup `13/13` GREEN.

Gate:

- [x] Architecture v4.11, Plan 022, and controls v1 define zero-quantity active
  creation before implementation;
- [x] focused RED/GREEN and packaged public-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation creates a zero-quantity disposable item and
  shows it in the managed list, Viewer, Receiving, and Production picker after
  Refresh.

### Slice 4an -- Mixed-UOM Process assembly retrieval

The operator's bottled chai Process combines one LB material requirement with
EA bottle and cap requirements and produces one EA bottled output. Retrieval
currently rejects that valid assembly because the historical Slice 4y contract
requires one common INPUT UOM. This is a deliberate D15 contract change: unlike
UOMs are not added or converted, but they may coexist as independently scaled
groups in one executable Process.

Required behavior:

- [x] Architecture v4.11, this plan, and controls v1 define mixed-UOM grouping
  before the conflicting retrieval behavior changes;
- [x] Process worksheet INPUT formulas group rows by normalized UOM, with one
  Batch basis quantity subtotal and one percentage total per UOM group;
- [x] every populated UOM group must total `100.0%` independently within the
  existing tolerance, while blank/catalog-invalid UOM and invalid quantities
  remain rejected;
- [x] **Retrieve Selected Process** accepts one Process containing LB material,
  EA packaging, and an EA output, imports all exact row quantities/UOMs through
  the real form handler, and deletes only the successfully retrieved table;
- [x] no implicit conversion is created: Recipe connections still require UOM
  compatibility and an actual conversion remains its own explicit Process; and
- [x] batch scaling, exact-key consumption, output creation, persistence,
  multi-table/Ctrl+click import, failed-table preservation, and prior Release 1
  GREEN behavior remain unchanged.

D13 RED sequence:

1. [x] Add a focused Slice 4an source contract and change the packaged public
   Process worksheet action evidence to require mixed-UOM retrieval; record RED
   while the actual handler returns **one compatible UOM** and preserves the
   selected table.
2. [x] Implement per-UOM validation in the worksheet controller without adding
   a global conversion or alternate authority.
3. [x] Rebuild the five-package set and rerun focused, packaged Production,
   Ribbon/compile, live-role, Release 1, NAS, static, and reviewed-growth gates.

Gate:

- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation retrieves `invSys_Process_004_20260827`
  with its LB and EA requirements and can save/release the resulting Process.

Automated evidence recorded 2026-08-29: focused source RED `1/7`, then `4/7`
after public-handler expectations were installed; focused GREEN is `7/7` and
the superseded Slice 4y source contract remains `6/6`. Packaged public Production
RED was `0/2` with `MixedUomAccepted=False`, `MixedUomRowsPreserved=False`, and
the exact one-compatible-UOM message. Final packaged Production plus clean
restart is `2/2`, using 4.5 LB concentrate, one EA bottle, one EA cap, and one
EA finished output; it records both mixed-UOM fields true and selected-table-
only deletion. Packaged XLAM is `81/81`, Ribbon/compile `142/142`, live roles
`47/47`, ordered Release 1 `30/30`, dedicated NAS runtime `16/16`, deterministic
static `19/19`, and reviewed growth/cleanup `13/13`. Static metrics remain 154
components, 5,177 procedures, and 1,043 candidates. Visible acceptance remains
open.

### Slice 4ao -- Recipe connection output names and catalog UOM selection

The 2026-08-29 visible Recipe Designer checkpoint confirmed that connections
work, but the Output selector exposes generated ID `006` rather than the
operator-facing output name and the connection UOM editor is free text. This is
a deliberate D15 operator contract correction; graph persistence continues to
use the exact generated Output ID.

Required behavior:

- [x] Architecture v4.11, this plan, and controls v1 define the corrected
  operator contract before implementation;
- [x] Recipe Designer's Output dropdown displays the selected Process output's
  name while retaining its generated Output ID as hidden selection authority;
- [x] Connect and Update stage and persist the exact hidden Output ID, never the
  displayed name;
- [x] connection UOM is a non-free-text dropdown populated from Settings'
  current Recipe UOM Catalog and an existing selected connection restores its
  catalog UOM; and
- [x] Recipe graph validation, compatible-UOM enforcement, saved definitions,
  Process/Recipe lifecycle behavior, Production runs, and prior Release 1 GREEN
  regressions remain unchanged.

D13 RED sequence:

1. [x] Add a focused source contract and extend the packaged public Production
   form action to exercise the real node/output controls and Connect/Update
   handlers; record RED while Output displays its ID and UOM is `TextBox`.
2. [x] Implement a name-visible/ID-bound output combo and catalog-backed UOM
   combo without changing the Designs Domain connection schema.
3. [x] Rebuild the affected package set and rerun focused, packaged Production,
   Ribbon/compile, live-role, Release 1, NAS, static, and reviewed-growth gates.

Gate:

- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation selects output names and a catalog UOM in
  Recipe Designer and successfully saves/releases the multi-Process Recipe.

Automated evidence recorded 2026-08-29: focused source RED was `0/7` and
focused GREEN is `7/7`. The packaged public Production callback plus clean
restart is `2/2` and records `RecipeOutputNameVisible=True`,
`RecipeOutputIdPreserved=True`, `RecipeUomCatalog=True`, and
`RecipeConnectionUpdated=True` through the actual Connect and Update handlers;
Recipe save/release/obsolete and the two-batch run also remain GREEN. Packaged
XLAM is `81/81`, Ribbon/compile `142/142`, live roles `47/47`, ordered Release 1
`30/30`, dedicated NAS runtime `16/16`, deterministic static `19/19`, and
reviewed growth/cleanup `13/13`. Static metrics are 154 components, 5,180
procedures, and 1,044 candidates. Visible acceptance remains open.

### Slice 4ap -- Named Recipe graph layout and finished-output guidance

The 2026-08-29 multi-Process Recipe checkpoint found that the graph technically
connects but still exposes generated Process/Output/Requirement codes in its
node and connection projections, supplies no column headers, and confines the
Connections list beside Recipe Process Nodes. The operator also expected a
terminal selection for the last output. Under D15, that last output is
deliberately left unconnected and becomes finished/co-product inventory. This
slice makes that contract visible without changing graph persistence.

Required behavior:

- [x] Architecture v4.11, this plan, and controls v1 define the named graph and
  finished-output contract before implementation;
- [x] From/To Process, Output, and Requirement selectors display names while
  retaining exact generated IDs as hidden bound values;
- [x] the editor labels its direction as **Upstream Process / Output /
  Downstream Process / Input Requirement / Qty / % / UOM** and rejects a
  same-node connection rather than suggesting that the last Process references
  itself;
- [x] Recipe Process Nodes and Connections have aligned column headers and use
  names instead of three-character Process/Output/Requirement codes;
- [x] Released Processes and Recipe Process Nodes sit above a full-width
  Connections list so names and allocation values remain readable;
- [x] Recipe Designer states that a final/co-product output stays unconnected
  and becomes managed inventory; and
- [x] Connect/Update/Disconnect, execution order, graph validation, save,
  release, restart, and Production runs preserve exact hidden identities and
  prior Release 1 behavior.

D13 RED sequence:

1. [x] Add a focused Slice 4ap source/layout contract and extend the packaged
   public Production action through the real named selectors, visible
   Connections selection, Connect, Update, and Disconnect handlers; record RED
   against the current ID-visible headerless side-by-side layout.
2. [x] Implement the name projections, headers, full-width layout, and explicit
   finished-output guidance while retaining the existing internal graph keys
   and same-node rejection.
3. [x] Rebuild and rerun focused, packaged Production, layout, Ribbon/compile,
   live-role, Release 1, NAS, static, and reviewed-growth gates.

Gate:

- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation builds and releases the multi-Process
  Recipe with its last output intentionally left as finished inventory.

Automated evidence on 2026-08-30 is GREEN. The focused source contract recorded
`2/9` RED and `9/9` GREEN. The packaged Production public action plus clean
restart is `2/2`, including named node/requirement/connection projections,
aligned headers, full-width geometry, visible selection, Connect/Update/
Disconnect, explicit same-node rejection, and finished-output guidance. The
packaged XLAM gate is `81/81`, Ribbon/VBA compile `142/142`, live roles `47/47`,
ordered Release 1 `30/30`, dedicated NAS runtime `16/16`, deterministic static
`19/19`, and reviewed growth/cleanup `13/13`. Static metrics are 154 components,
5,186 procedures, and 1,045 candidates. Visible acceptance remains open.

### Slice 4aq -- Output-first Recipe routing and fork/convergence flow

The 2026-08-30 visible checkpoint showed that Slice 4ap made graph identities
readable but still presented the persistence edge editor. Selecting Brewed
Black Slurry Tea defaulted **Downstream Process** to Chai Spice Mix and exposed
Chai Spice Mix ingredient requirements, even though both Processes should run
independently and their outputs should converge on Black Scottie Chai Classic
Concentrate. This is a deliberate operator-contract correction under D15; the
stored output-to-requirement graph remains authoritative.

Required behavior:

- [x] Architecture v4.11, this plan, and controls v1 define output-first
  compatible routing, fork/convergence projection, and derived stages before
  implementation;
- [x] selecting a Process output filters **Feeds Process** to Recipe nodes with
  exactly one compatible unresolved requirement, based on output item/SKU,
  Ingredients Assignment, and UOM;
- [x] the compatible `RequirementId` is bound internally; Recipe Designer does
  not expose an arbitrary ingredient dropdown or automatically choose the next
  Recipe node;
- [x] the full-width **Output Flow** projection shows Stage, Produced by,
  Output, Feeds Process, Qty, %, and UOM, including terminal **Finished
  inventory** output rows;
- [x] two independent first-stage outputs can converge on separate requirements
  of one later Process, whose output can feed the next stage;
- [x] Auto Order derives visible topological stages while retaining unique
  deterministic persisted execution ordinals; and
- [x] Connect/Update/Disconnect, graph validation, lifecycle, restart, and
  Production Run preserve hidden IDs and existing Release 1 behavior.

D13 RED sequence:

1. [x] Add a focused Slice 4aq source/layout contract and extend the packaged
   public Production action through output selection, compatible Feeds Process,
   Connect, Auto Order, visible Output Flow selection, and terminal-output
   projection; record RED against the arbitrary downstream/ingredient editor.
2. [x] Implement compatibility filtering, internal requirement binding,
   fork/convergence stage projection, and terminal output rows without changing
   Designs Domain graph schema or headless authority.
3. [x] Rebuild and rerun focused, packaged Production, layout, Ribbon/compile,
   live-role, Release 1, NAS, static, and reviewed-growth gates.

Gate:

- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation builds/releases the sample fork:
  Brewed Black Slurry Tea + Chai Spice Mix -> Black Scottie Chai Classic
  Concentrate -> Bottle 64oz Chai Classic -> Finished inventory.

Automated evidence (2026-08-30): focused Slice 4aq `2/8` initial contract
presence and six behavioral REDs, then `8/8` GREEN; Slice 4ap `9/9`, Slice 4ao
`7/7`, Slice 4an `7/7`, and Production layout `8/8`; packaged Production
action plus clean restart `2/2`; packaged XLAM `81/81`; Ribbon/compile
`142/142`; live roles `47/47`; ordered Release 1 `30/30`; dedicated NAS
runtime `16/16`; deterministic static `19/19`; reviewed growth/cleanup
`13/13`. Static metrics are 154 components, 5,201 procedures, and 1,048
candidates. Visible sample-Recipe acceptance remains open.

### Slice 4ar -- Output-yield clarity, list headers, and multi-Process run plan

The 2026-08-30 visible checkpoint accepted the fork/convergence direction but
found three blocking projection defects. Recipe Designer showed a downstream
requirement's consumption quantity and percentage as though they were the
producing Process output yield; Process Designer could display blank output
yield defaults and appeared to clear them on Update; and Ingredients
Assignment plus Production Run did not provide enough named, header-backed
context for an operator to understand a multi-Process Recipe. This slice is a
deliberate D15 operator-contract correction; persisted graph identities,
released quantities, exact-key allocation, and completion authority remain
unchanged.

Required behavior:

- [x] Architecture v4.11, this plan, and controls v1 distinguish producing
  **Output Qty / Yield %** from downstream **Required Qty / Required %** before
  implementation;
- [x] Output Flow always projects the producing output definition, including
  75.025 LB / 100% for a quantity-defined Chai Spice Mix output, while the
  connection editor retains the downstream 9.5 LB requirement separately;
- [x] quantity-defined outputs normalize blank yield fields to 100% and their
  output Qty as Yield basis quantity, while explicit percentage/basis values
  survive Update, save, and reload;
- [x] Saved Processes, Requirements, Outputs, Instructions, Processes,
  Ingredient Requirements, Managed Items, and Acceptable Items all have
  aligned column headers;
- [x] Acceptable Items shows managed item name, UOM, and item/SKU code while
  keeping Requirement ID internal;
- [x] Production Run - List shows the whole released multi-Process plan and
  identifies each external inventory choice by its owning Process name; and
- [x] a disabled **Scale from target output Qty (coming later)** option is
  visible but performs no calculation or state change in this slice.

D13 RED sequence:

1. [x] Add a focused Slice 4ar contract and extend the packaged public
   Production action through the real output Update, Recipe Output Flow,
   Ingredients Assignment, and run-load handlers; record RED against the
   ambiguous/blank/headerless projections.
2. [x] Implement only the approved projections/default normalization and
   multi-Process layout, preserving Domain/Core headless authority and hidden
   identities.
3. [x] Rebuild and rerun focused, packaged Production, layout, Ribbon/compile,
   live-role, Release 1, NAS, static, and reviewed-growth gates.

Gate:

- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation verifies the sample four-Process Recipe,
  named acceptable items, output yields, and multi-Process run plan.

Automated evidence recorded 2026-08-30: focused Slice 4ar RED `0/8` then
GREEN `8/8`; prior Output Flow contract `8/8`; Production layout `8/8`;
packaged public Production actions `2/2`; packaged XLAM regression `81/81`;
Ribbon/VBA compile `142/142`; live roles `47/47`; ordered Release 1 `30/30`;
dedicated NAS `16/16`; deterministic static `19/19`; and reviewed growth
`13/13`. The remaining gate is visible operator acceptance.

### Slice 4as -- Compact Process Output editor and catalog UOM

The 2026-08-30 visible Slice 4ar checkpoint found that Process Designer's
hidden output SKU text box still reserved a blank slot between Output and
Design, and that Output UOM appeared as a free-text box on a second row. This
is a newly discovered visible-acceptance blocker. It changes only the
operator-facing Output editor geometry and UOM selection contract; Process
identity, output SKU/design identity, quantities/yields, lifecycle authority,
and worksheet behavior remain unchanged.

Required behavior:

- [x] Architecture v4.11, this plan, and controls v1 require one compact Output
  editor row and a Recipe UOM Catalog dropdown before implementation;
- [x] the hidden output SKU control reserves no visible horizontal gap;
- [x] Design, Version, Output Qty, Yield %, and Yield basis move left into the
  recovered space, and UOM appears on that same row at the right;
- [x] Output UOM is a dropdown-list populated from Settings' current Recipe UOM
  Catalog and restores the saved selection on load; and
- [x] Add/Update continue through the existing public form handlers and reject
  an Output UOM that is not in the current catalog.

D13 RED sequence:

1. [x] Add a focused Slice 4as source contract and extend the packaged public
   reusable-Production action through the real Output Add/Update handlers;
   record RED against the gap, second-row text box, and missing catalog proof.
2. [x] Implement only the compact geometry and catalog-backed Output UOM
   projection, preserving hidden IDs and Designs Domain authority.
3. [x] Rebuild Operations and rerun focused, packaged Production, layout,
   Ribbon/compile, Release 1, NAS, static, and reviewed-growth gates.

Gate:

- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation shows one compact Output editor row with a
  working catalog UOM dropdown.

Automated evidence recorded 2026-08-30: focused Slice 4as RED `1/6` then
GREEN `6/6`; prior Slice 4ar `8/8`; Production layout `8/8`; packaged public
Production actions and clean restart `2/2`; packaged XLAM `81/81`;
Ribbon/VBA compile `142/142`; live roles `47/47`; ordered Release 1 `30/30`;
dedicated NAS `16/16`; deterministic static `19/19`; and reviewed growth
`13/13`. Static metrics are 154 components, 5,210 procedures, and 1,048
candidates. Visible Process Designer acceptance remains open.

### Slice 4at -- Location-stock Production allocation and Receiving capacity stub

The 2026-08-30 visible checkpoint confirmed that Ingredients Assignment finds
one Cassia Oil managed stock at CLEARVIEW, but Production Run exposed four
receipt entities as four operator choices. Architecture v4.11 retains those
exact keys for audit and consumption, while this deliberate D14/D15 projection
correction makes the operator choose warehouse stock by managed item and
location. The same checkpoint found Ingredients Assignment's 45-point
`System_Key` column unreadable and requested a future Receiving location-
capacity concept.

Required behavior:

- [x] Architecture v4.11, this plan, and controls v1 distinguish a visible
  location-stock bucket from its hidden contributing exact entities before
  implementation;
- [x] Ingredients Assignment gives `System_Key` enough width to read the
  selected managed entity without changing the saved SKU alternative;
- [x] Production Run - List shows one acceptable stock row per managed
  SKU/UOM/Location/Condition bucket with summed available stock, not one row
  per Receiving entry;
- [x] applying a bucket quantity through the existing Apply handler expands it
  deterministically across sufficient exact keys, and Check In/completion keep
  exact-key stale/overdraw/location validation;
- [x] Utility/Service stock remains non-counted and displays its existing
  nonnumeric quantity-mode label; and
- [x] Receive Item Results shows a blank **Capacity (coming later)** column
  that has no validation, persistence, or inventory effect in this slice.

D13 RED sequence:

1. [x] Add a focused Slice 4at contract and extend the packaged Production and
   Receiving public form actions; record RED against receipt-row choices, the
   narrow key, missing bucket expansion, and missing Capacity stub.
2. [x] Implement only the stock projection/allocation adapter, readable key
   width, and inert Capacity column; preserve exact Domain entities and event
   payload keys.
3. [ ] Rebuild Operations and rerun focused, packaged Production/Receiving,
   layout, Ribbon/compile, live-role, Release 1, NAS, static, and reviewed-
   growth gates.

Gate:

- [x] focused RED/GREEN and packaged real-handler evidence are recorded;
- [x] applicable local and isolated packaged regressions remain GREEN; and
- [ ] visible operator confirmation shows one Cassia Oil/CLEARVIEW stock row,
  readable assignment identity, and the inert Receiving Capacity column.

Automated evidence recorded 2026-08-30: focused RED `0/7` then GREEN `7/7`;
packaged Production public actions, two-batch run, and clean restart `2/2` with
`LocationStockBuckets=True` and `LocationStockExactExpansion=True`; packaged
Receiving durability `1/1` with `CapacityStub=True`; prior Production source
regressions `6/6`, `7/7`, `8/8`, and `6/6`; Receiving stabilization `10/10`;
workflow readiness `18/18`; launcher contracts `24/24`; packaged XLAM `81/81`;
Ribbon/VBA compile `142/142`; live roles `47/47`; ordered Release 1 `30/30`;
deterministic static `19/19`; and reviewed growth `13/13`. Static metrics are
154 components, 5,223 procedures, and 1,050 scanner candidates. The current
dedicated-NAS rerun remains unchecked because its configured test root was
unavailable; the last verified prior NAS result remains `16/16`. Visible
operator confirmation remains open.

### Slice 4au -- Receiving 10-column results and non-wrapping headers

The 2026-08-30 visible Receiving checkpoint exposed a runtime error 380 while
searching Receive Item Results. Slice 4at had placed the hidden representative
`System_Key` plus ten operator fields into an eleven-column MSForms ListBox,
but the control's `.List` projection supports only ten columns. The same
checkpoint showed header labels wrapping over Received Tally and Aggregate
Received. Production's corrected one-row-per-location stock projection was
visibly accepted and remains unchanged.

Required behavior:

- [x] Receive Item Results contains exactly ten visible columns: Code, Item,
  UOM, Available, Location, blank **Capacity (coming later)**, Lot, Condition,
  Description, and Vendor;
- [x] the representative exact `System_Key` remains hidden in an internal row
  map aligned with the visible results and is used by the real Add Selected and
  disposition actions;
- [x] a non-empty search/change action repopulates the result list without an
  invalid List property write or runtime error 380;
- [x] Receiving, history, tally, and aggregate column-header labels are fixed
  to one line and cannot wrap over their ListBoxes; and
- [x] the blank Capacity stub remains inert and Production allocation remains
  the accepted location-stock behavior from Slice 4at.

D13 RED sequence:

1. [x] Record focused RED against the eleven-column projection, missing hidden
   row map, non-empty search evidence, wrapping headers, and missing packaged
   evidence flags.
2. [x] Move hidden identity out of the ListBox, remap all selection consumers,
   and enforce single-line header labels through the existing form actions.
3. [x] Rebuild Operations and rerun focused, packaged Receiving, compile,
   Release 1, static, and reviewed-growth gates.

Gate:

- [x] focused RED is recorded as `0/7`;
- [x] focused GREEN and packaged non-empty search evidence are recorded;
- [x] applicable regressions remain GREEN; and
- [ ] visible operator confirmation shows searchable inventory results and
  readable one-line tally/aggregate headers.

Automated evidence recorded 2026-08-30: focused RED `0/7` then GREEN `7/7`;
packaged Receiving durability/search/header action `1/1` with
`SearchRowsLoaded=True`, `HiddenSystemKeyMap=True`,
`TenColumnItemResults=True`, and `HeadersSingleLine=True`; Receiving
stabilization `10/10`; prior Slice 4at `7/7`; packaged XLAM `81/81`;
Ribbon/VBA compile `142/142`; live roles `47/47` after one transient Excel COM
automation retry; ordered Release 1 `30/30`; deterministic static `19/19`;
and reviewed growth `13/13`. Static metrics are 154 components, 5,224
procedures, and 1,050 scanner candidates. The current dedicated-NAS rerun was
not repeated because its configured test root remains unavailable; the prior
verified result remains `16/16`.

### Slice 4av -- Process-scoped Production execution, plan sufficiency, and run instructions

The 2026-08-30 visible four-Process Recipe checkpoint selected Chai Spice Mix,
allocated its four ingredients, and then found Check In blocked by unresolved
Dry Crushed Black Tea belonging to Brewed Black Slurry Tea. The checkpoint also
showed only four inventory rows and no execution instructions. This is a
deliberate D15 contract change from atomic whole-Recipe completion to one
selected Process at a time; Architecture v4.11 and controls v1 define the new
boundary before implementation.

Required behavior:

- [x] Architecture v4.11, this plan, and controls v1 define Process-scoped
  Check In/completion and whole-Recipe status before VBA changes;
- [x] selecting a Process filters its allocation surface and Check In validates
  only that Process's external inputs and completed upstream commitments;
- [x] unresolved or insufficient inventory for another Process remains visible
  in Multi-Process Run Plan but does not block the selected READY Process;
- [x] the plan displays READY, NEEDS ALLOCATION, WAITING UPSTREAM,
  ! INSUFFICIENT, or COMPLETE status using the complete Recipe context;
- [x] independent READY Processes may run in either order, while downstream
  Processes cannot run before sufficient exact upstream output exists;
- [x] completing one Process consumes only its inputs and creates each of its
  outputs under a new exact `System_Key`; the batch completes only when every
  Recipe Process is complete;
- [x] Acceptable Inventory For Run shows at least eight ordinary rows at the
  default size; and
- [x] the selected released Process's ordered instructions are exposed read-
  only on Production Run - List.

D13 RED sequence:

1. [x] Add a focused Slice 4av contract and extend the packaged reusable-
   Production public action through Process-filter change, Check In, Complete
   Run, and run refresh; record behavioral RED against Recipe-global blocking,
   missing status/instructions, and four-row palette geometry.
2. [x] Implement Process-scoped session state, dependency validation, exact
   intermediate continuity, status projection, instructions, and layout while
   preserving headless Domain/Core authority and existing exact-key events.
3. [ ] Rebuild Operations and rerun focused, packaged Production, layout,
   Ribbon/compile, live-role, Release 1, static, and reviewed-growth gates.

Gate:

- [x] focused RED/GREEN and packaged same-handler evidence are recorded;
- [x] applicable source, compile, layout, and static regressions remain GREEN;
- [ ] the complete packaged ProductionReusable sequence reruns without the
  host-level `OFFICE_MODULE_VERSION_MISMATCH` crash observed twice in the
  pre-existing Process worksheet bulk-import stress test; and
- [ ] visible confirmation completes Chai Spice Mix while unrelated Dry Tea is
  unresolved, shows eight palette rows/instructions, and then permits the next
  dependency-ready Process.

Automated evidence recorded 2026-08-30: focused source initial RED (1 passed,
7 RED), then GREEN `8/8`; focused packaged public action `1/1` with
`SelectedProcessOnly=True`, `RunInstructionsVisible=True`,
`WholeRecipeStatus=True`, `EightPaletteRows=True`, two completed batches,
exact input/output keys, routed-intermediate consumption, and retained
co-product balance; Operations shadow compile/initialization `13/13`; packaged
Production layout GREEN at minimum/default/expanded and native window states;
maintained Production source regressions GREEN, including reusable Production
`10/10`; deterministic static/reviewed-growth `19/19` at 154 components, 5,249
procedures, 1,050 scanner candidates, and 1,052 reviewed candidates. The full
ProductionReusable sequence was attempted twice but Excel crashed during the
older Process worksheet bulk-import test before the new run action, with
Windows recording `OFFICE_MODULE_VERSION_MISMATCH`; the focused packaged run
then completed independently and GREEN. Visible operator confirmation remains
open.

### Slice 4aw -- Released Process editing and worksheet export

The 2026-08-30 visible checkpoint completed the first three Processes in the
four-Process Chai Recipe, then found that the released mixed-UOM bottling
Process could be loaded but not saved after its requirement/output quantities
were changed. The same checkpoint requested an explicit way to send that
already-made Process to a worksheet table, edit it, and retrieve it. This is a
correction to the existing Architecture v4.11 D15 immutable-version and Process
worksheet workbench contract; it does not authorize rewriting a saved version.

Required behavior:

- [x] Architecture v4.11 remains authoritative: a saved or released Process is
  immutable, and editing creates the next available generated DRAFT version;
- [x] Process Designer labels the read-only reference action **View Process**
  and the editable successor action **Edit as New Version** so operators do not
  mistake a loaded immutable version for a writable draft;
- [x] **Edit as New Version** loads the complete selected definition, proposes
  the next version, and rebases every retained Output Design Version to that
  successor while preserving Process, Requirement, Output, instruction, UOM,
  managed SKU, and Ingredient Assignment identity;
- [x] the worksheet action is labelled **Send Process to Sheet** and sends the
  current new/editable Process—or a viewed saved Process promoted to its next
  version—to a uniquely named table in the captured
  `Production.Operator.xlsm`;
- [x] worksheet retrieval uses the existing public DRAFT-save authority,
  removes only a successfully saved selected table, and leaves Release as a
  separate explicit action; and
- [x] the corrected workflow accepts mixed-UOM assembly quantities such as
  concentrate in LB plus bottle/cap/output in EA without conversion or
  immutable-version overwrite.

D13 RED sequence:

1. [x] Add focused Slice 4aw source checks and record RED for ambiguous action
   wording, missing successor Output Design Version rebasing, and absent
   packaged edit/export evidence (`0/6`).
2. [x] Extend the packaged public Production action through selected released
   Process view, edit-as-successor, row Update, Send Process to Sheet, table
   edit, Retrieve Selected Process, and resulting DRAFT verification.
3. [x] Rebuild Operations and rerun focused, packaged Production worksheet/
   lifecycle, compile/layout, Release 1, static, and reviewed-growth gates.

Gate:

- [x] focused RED/GREEN and packaged same-handler evidence are recorded with
  `ReleasedProcessEditable=True`, `ExistingProcessExported=True`,
  `ExportRoundTrip=True`, and `OutputDesignVersionRebased=True`;
- [x] existing mixed-UOM worksheet, reusable Production, and launcher
  regressions remain GREEN; and
- [ ] visible acceptance edits/releases the corrected bottling Process and
  completes it after its dependency-ready concentrate Process.

Automated evidence recorded 2026-08-30: focused source began RED at `0/6` and
is GREEN at `7/7`. The focused packaged public action is GREEN at `1/1` with
`ReleasedProcessEditable=True`, `ExistingProcessExported=True`,
`ExportRoundTrip=True`, `OutputDesignVersionRebased=True`, and
`OutputYieldRebased=True`. Packaged XLAM validation is `81/81`, packaged
Ribbon/VBA compile validation is `142/142`, Production layout validation passes
at minimum/default/expanded/native sizes, Slice 0 tooling is `62/62`, the full
Release 1 chain is `30/30`, and live role workflows are `47/47`. Static
maintenance evidence records 154 components, 5,253 procedures, 1,050
candidates, and 1,052 reviewed entries. The older combined reusable-Production
stress sequence still encounters the recorded Office
`OFFICE_MODULE_VERSION_MISMATCH` crash at its pre-existing output-picker step;
the isolated same-handler Slice 4aw proof and the Release 1 regressions are
GREEN. Visible correction and completion of the operator's bottling Process
remain open.

### Slice 4ax -- Routed intermediate visibility and full multi-Process acceptance: user accepted

The 2026-08-30 visible checkpoint selected the final mixed-UOM bottling Process
and found that its connected upstream concentrate was absent from **Acceptable
Inventory For Run** while the packaging inputs were present. Source inspection
shows that the palette intentionally excludes incoming Recipe connections: it
is an allocation surface for selectable external stock, while a routed
intermediate must come from the exact output key created by the upstream
Process in the same active run. The screenshot showed **WAITING UPSTREAM**,
blank active output keys, and zero Process totals, so it does not establish that
the earlier Process completions belonged to the currently loaded run.

The same inspection exposed a genuine operator-visibility gap. The completion
service includes routed intermediate keys in its `USED` event items, but the
Inventory Check projection enumerates only external allocation records. The
single numeric **Used Goods** summary also cannot communicate consumption from
more than one UOM group.

Required next behavior and proof:

- [x] establish through the public form handlers whether the completed upstream
  Processes and the selected downstream Process remain in one active run across
  the exact operator navigation/refresh sequence;
- [x] keep incoming Recipe connections out of the selectable external-stock
  palette, but expose every ready routed intermediate as a read-only exact-key,
  quantity, UOM, and source-Process row in Inventory Check or an equally clear
  operator projection;
- [x] prove that downstream Check In changes from **WAITING UPSTREAM** to
  **READY** only after the upstream Process completes in the same run;
- [x] prove downstream completion queues and applies the routed output's exact
  key as `USED`, leaves any uncommitted balance as managed inventory, and makes
  that consumption visible in Events/Viewer;
- [x] reconcile the **Used Goods** display contract for mixed-UOM Processes
  before changing it; a single unitless total must not add unlike UOMs; and
- [ ] complete one visible four-Process fork/convergence Recipe from initial
  allocation through final bottling without clearing, replacing, or silently
  restarting the active run.

Approved 2026-08-30 contract refinement:

- **Used Goods** is a read-only, deterministic UOM-grouped summary for all
  scaled goods consumed by one Process/batch, such as `5 LB; 12 EA`; it does
  not add unlike units and repeats consistently on that Process's output rows.
- `EA`, case-insensitively after UOM normalization, is whole-unit only across
  definition, staging, allocation, event, and Inventory Domain application
  boundaries. Fractional EA is rejected without rounding. This is a cross-role
  invariant, so its D13 proof must include the relevant packaged public handler
  and the Inventory Domain apply guard.

D13 next test: extend the packaged reusable-Production public action through
upstream completion, refresh/navigation, downstream selection, routed-key
Inventory Check visibility, exact-key consumption, remaining balance, and
multi-UOM output history. Record meaningful behavioral RED before changing the
form or run-session implementation.

Automated evidence recorded 2026-08-31: focused Slice 4ax source began at
`5/7` with two named-Chai handler assertions RED, then reached `7/7` GREEN.
The packaged public Production action created and released the four
Process **Classic Chai Fork Convergence** graph, completed Tea Brewing and
Spice Blending, proved both exact routed keys in Convergence Inventory Check,
consumed both keys, completed Final Bottling from the exact concentrate key,
and created the final bottled `EA` output under a new key. It recorded
`ChaiInitialWaitingUpstream=True`, `ChaiRoutedConvergenceInputs=True`,
`ChaiUpstreamExactKeysConsumed=True`, `ChaiFinalBottlingRoutedInput=True`,
`ChaiFinalOutputNewKey=True`, `ChaiFourProcessesCompleted=True`,
`ChaiFinalBottlingCompleted=True`, and `ChaiRunNotRestarted=True`. The test
uses the packaged form handlers and the same visible Inventory Check and
Production Output projections as an operator. This is automated evidence only;
the dedicated NAS human-visible UAT checkbox remains open.

The current approved five-package set was also revalidated on 2026-08-31
against the dedicated NAS test leaf in two clean Excel sessions: 16/16 launcher,
package-hash, target-selection, and read-only-safety checks passed with zero
canonical-file changes from launcher use. That deployment check does not create
or complete the Chai Recipe and is not human-visible Chai acceptance.

Visible NAS Chai UAT record (open): use only the dedicated test warehouse and
the approved package set. Load one released four-Process Chai Recipe and note
the displayed Recipe version, Run ID, and Batch before any Process completes.
Before upstream completion, select Convergence and confirm its connected inputs
read **WAITING UPSTREAM** and that **Acceptable Inventory For Run** contains
only selectable external stock. Complete Tea Brewing and Spice Blending through
the public **Check In** and **Complete Run** actions; use normal **Refresh** and
Process selection to revisit Convergence without clicking **New**, reloading a
Recipe, or starting a new batch. Confirm the same Recipe/version/Run/Batch is
still displayed and that Inventory Check now has one read-only row for each
routed input, identifying source Process/output, downstream requirement, exact
`System_Key`, committed quantity/UOM, and remaining balance. Complete
Convergence, then use the same Refresh/navigation sequence before Final
Bottling and confirm its exact concentrate key appears as a read-only routed
input. Complete Final Bottling and confirm all four Process statuses are
complete, the output has a new `System_Key`, any co-product balance remains
visible, **Used Goods** keeps LB and EA in separate groups, and Viewer Events
shows correlated Production Input Consumed/Production Output Created entries.
Return redacted screenshots of the initial waiting state, both routed-input
checks, final output/co-product state, Viewer Events, and the unchanged
Recipe/version/Run/Batch context.

### Slice 4ay Production Run - List responsive readability: user accepted

The user reported that **Inventory Check** and **Selected Process
Instructions** were too short, that vertical form resize gave all additional
space to **Production Output**, and that right-dragging moved the Instructions
list's left edge. This is a visible Operations layout refinement, not a change
to D15 Production execution, `System_Key`, run persistence, allocation, or
completion semantics.

Required behavior:

- [x] reserve at least eight rendered rows for **Inventory Check** and four for
  **Selected Process Instructions** at the readable `1110 x 800` point
  baseline;
- [x] make Recipes, **Multi-Process Run Plan**, **Acceptable Inventory For
  Run**, **Inventory Check**, **Selected Process Instructions**, and
  **Production Output** share added/removed vertical space proportionally;
- [x] keep **Selected Process Instructions** fixed to the left while its right
  edge follows horizontal form resize; and
- [x] move the list captions and column headers with their matching list bands
  so audit columns remain visually attributable during resize.
- [x] keep a vertical page scrollbar visible on **Production Run - List** so
  constrained-height operators can reach every list and action; and
- [x] keep the Run Plan **Qty** header in its own header band during vertical
  resize, rather than treating it as the palette's `Qty` caption.

D13 evidence: the new focused source contract was meaningfully RED at `0/5`
before implementation and GREEN at `5/5` after it. The change extends the
Operations-local Andy Pope anchor implementation with declarative proportional
vertical control registration; the form resize callback still contains no
one-off coordinate arithmetic. The public
`mProduction.BtnOpenProductionForm` packaged proof is GREEN with
`CheckEightRows=True`, `InstructionsFourRows=True`, `AllListsGrew=True`,
`InstructionLeftStable=True`, `HeadersAligned=True`, and
`GeometryHealthy=True`; existing reusable two-batch and four-Process Chai
handler proofs also remain GREEN. The native Windows API validator passed
minimum/default/expanded, minimize/restore/maximize, bounds, overlap, and
full-client-fill checks using Production Run - List screenshots. A final
operator drag-resize retest of the deployed package remains open before visible
acceptance is checked complete.

The two follow-on checks were meaningfully RED at `5/7` and GREEN at `7/7`.
The public launcher additionally reports `RunListVerticalScrollbar=True` and
`RunPlanQtyHeaderAligned=True`; the latter proves the `Qty` header remains
co-band with the other Run Plan headers after vertical resize. Native default
and expanded captures visibly show both the vertical scrollbar and the retained
header. This is presentation-only and retains the earlier two-batch and Chai
handler GREEN results.

### Slice 4az Actual-output regulation and tolerance: user accepted

The observed Chai run correctly rejected `605 LB` where the routed downstream
commitment is `610 LB`; that exact-key sufficiency rule remains mandatory.
This slice adds no input-to-output mass-balance model. Planned output remains a
scaled comparison value while the operator-entered **Actual Output** remains
the created inventory quantity and the only physical measurement authority.

Approved behavior:

- [ ] add a sixth **Production Settings** page to the packaged Production form;
  it edits draft versioned definitions, never an active run or released version;
- [ ] support an output-level Process default and an optional pinned
  Recipe-node/output override, each with **Regulated**, **Floor**, and
  **Ceiling** in the output's catalog UOM;
- [ ] scale enabled bounds with the batch and require actual output in that
  interval, while retaining exact route sufficiency with effective lower bound
  `max(scaled Floor, routed commitment)`;
- [ ] allow a `600–610 LB` regulation range with a `610 LB` route: its effective
  lower bound is 610; reject only a ceiling below the routed commitment;
- [ ] retain positive actuals, exact `System_Key` output creation, whole-EA
  validation, routed consumption, co-product balances, and correlated events;
  and
- [ ] leave unregulated outputs under the existing positive-actual plus routed
  commitment rule.

The visible **Production Settings** page must also include these operator
instructions: Process defaults are edited only while a Process is DRAFT in
**Process Designer**; Recipe overrides are edited only while a Recipe is DRAFT
in **Recipe Designer** after its Process nodes are present; each change then
requires that respective Draft Save and Release. **Production Run** only reads
the released Recipe version and cannot modify an active run, a released
Process, or a released Recipe. Recipe-node selection must persist while the
operator applies or inspects an override rather than resetting to the first
node.

D13 scope before implementation: a public `mProduction.BtnOpenProductionForm`
test must be meaningfully RED for draft save/release/reopen of Process defaults
and Recipe overrides; regulated below-floor, above-ceiling, and below-route
rejection; valid `610 LB` completion for the `600–610 LB`/`610 LB` case;
unregulated compatibility; fractional-EA rejection; exact routed-key
consumption; and no input/output mass-balance comparison. Package scope is
Operations plus Designs Domain schema/projection authority. A user actively
rebuilding a Recipe must first release a new version; an existing active run is
never altered, cleared, or restarted by this feature.

### Slice 4ba -- variable Process quantity modes: user accepted

Certain Process outputs are indeterminate until production measures them. Add
versioned `OutputQtyMode` with `FIXED` (current behavior) and `ACTUAL`
(variable, determined by Actual Output). Process Designer will expose the two
choices as **Enter a number** and **Variable -- determined by Actual Output**;
the existing numeric Output Qty entry remains active only for the former.

For `ACTUAL`, Qty, Yield %, and Yield basis are blank, Actual Output is the
required positive completion measurement, and that exact quantity creates the
new output `System_Key`. Whole-EA enforcement remains unchanged. An ACTUAL
output may route to fixed-quantity requirements; recipe release permits no
percentage connection from it, and execution requires Actual Output to meet
the batch-scaled routed commitment before downstream exact-key consumption.
Regulation remains optional and route-safe; it is not an input/output mass
balance calculation. Process-to-Sheet will add a validated **Qty Mode** column
and must preserve the mode through Send/Retrieve. Existing definitions without
the new field read as `FIXED` and are never rewritten.

The expanded proposal also adds `RequirementQtyMode` with `FIXED` and
`ACTUAL`. A variable input is shown as **Variable -- determined at Check In**:
it has no planned Qty, Percent, or Batch basis; the operator selects compatible
external stock through the existing external-only palette and commits the
positive measured amount by exact `System_Key` at Check In. That committed
amount is the auditable input fact; Actual Output does not calculate it and no
mass balance is introduced. Whole-EA validation remains mandatory.

An `ACTUAL` input is external-only in Release 1. Recipe release rejects an
incoming routed connection to it, which preserves both the current read-only
routed-input display and the non-negotiable exact routed commitment. The
already-proposed `ACTUAL` output can still supply a fixed downstream
requirement when its measured actual is sufficient. Variable routed input is a
separate future contract, not an implicit hybrid.

D13 scope after approval: public Process Designer Add/Update/Save/Release and
Send/Retrieve handlers must show RED then GREEN for fixed compatibility,
variable save/reopen, variable worksheet round trip, fractional-EA rejection,
rejection of variable numeric yield fields and percentage routes; variable
input external Check In/refresh/reopen; rejection of a routed variable input;
and a variable routed-output run that blocks below commitment then consumes the
exact actual-output `System_Key` at a sufficient actual. Package scope is
Operations and Designs Domain schema/query/projection authority.

Implementation now persists `RequirementQtyMode` and `OutputQtyMode`, presents
the two Process Designer mode dropdowns, and adds the Process worksheet **Qty
Mode** validation/round trip. External `ACTUAL` requirements accept a positive
measured exact-key allocation at Check In without a planned-quantity cap;
routed `ACTUAL` requirements remain release-rejected. Focused source and public
packaged form-action GREEN are recorded; the dedicated actual-input run and
worksheet round-trip acceptance cases remain required before slice completion.

### Slice 4bb -- external-stock UOM conversion: user accepted

The Citrus Oil Production checkpoint released a Recipe whose requirements and
output are `OZ`, while selected external stock remains physically held in `LB`.
The current plan/spec correctly preserves the stock's native UOM, but it has no
approved conversion path; showing it as an ordinary candidate is insufficient.

This approved contract follows Architecture v4.11 Slice 4bb. It
adds a versioned UOM Catalog workbench without another form: **Edit UOM Catalog
on Sheet** writes one captured-workbook fill-out table, and **Retrieve UOM
Catalog** validates the selected table then publishes a new catalog version.
The table defines `UOM`, Dimension, Base UOM, Units Per Base UOM, Convertible,
Enabled, and Notes. New mass/volume units automatically interconvert only when
their published Dimension/base entries are valid. The initial catalog declares
`LB`, `LBS`, `OZ`, `KG`, and `G` against MASS/LB and `GAL`, `QT`, `PT`, `L`, and
`ML` against VOLUME/GAL.

Conversion remains external Production allocation only. The palette shows
Requirement UOM, Stock UOM, native available, and converted available; the
operator enters Requirement UOM. The exact stock `System_Key` remains native,
and events retain both quantities/UOMs plus catalog-version/factor evidence.
`EA` is never convertible. `CS` is initially nonconvertible because its
case-to-each factor is item/package-specific, not globally safe. Routed Recipe
edges remain UOM-equal and a routed transformation still requires an explicit
Process.

Before implementation: create a focused public
`mProduction.BtnOpenProductionForm` RED using a fixed `OZ` requirement and `LB`
external key; then prove the published LB/OZ relationship, exact-key native
balance preservation across Refresh/reopen, event audit fields, UOM-workbench
Send/Retrieve of an added same-dimension unit, and rejection of missing,
disabled, cross-dimension, `EA`, and `CS` conversions. Expected packages:
Core/Inventory configuration authority and Operations Production; Designs
definitions are not rewritten.

### Slice 4bc -- Production batch notes and Viewer clarity/export: Production portion user accepted; Viewer remains in Release 1 scope

This user-approved Release 1 slice corrects the visible external-stock UOM
pairing and completes a bounded operator usability improvement without changing
inventory authority. In **Acceptable Inventory For Run**, both paired columns
use native-first order: **Stock / Requirement UOM** and **Native / Requirement
Available**. Their header geometry must remain readable at default/minimum and
responsive sizes.

Production Run - List adds one optional **Batch Note** for the active Recipe
version/RunId/batch. It is editable until that batch's first Check In, then
frozen and retained through Refresh and Process navigation. Each production
consume/create event for that batch preserves the same note as correlated audit
detail. Next Batch and Clear Run start a distinct blank note. It is not a
quantity calculation, route override, or authority-write shortcut.

Operations renames the operator-facing **Inventory Viewer** to **Viewer** and
moves **invSys Sign In** directly to the right of **Runtime Context**. Viewer
keeps its current Inventory and Events read-only projections, fixes Events
headers to calculated list-column geometry, and adds **ListBox->Table**. The
operator enters a declared currently open list-box name and clicks **Export
ListBox to Table** to write its already loaded visible rows/headings to a new
worksheet table. The action does not refresh, process, write NAS authority, or
expose hidden/internal columns. A declared Admin list resolves only for an
authenticated Administrator.

D13 scope: first observe focused RED through the same Production public Check
In/Complete handlers for note freeze/Refresh/event detail and through the
packaged Viewer/ribbon callbacks for label/order/export geometry. GREEN must
prove native-first UOM pairing, note retention and event correlation, an
Operations list export with headings/no hidden key, Admin-source capability
gating, Viewer event header/list alignment, and ribbon placement. Expected
package scope is Operations plus the existing Production event builder; Core
and Domain contracts remain headless and do not gain a Viewer write path.

### Slice 4bd -- superseded single-server physical-UAT wording (see approved 2026-09-03 amendment below)

This slice performs the existing Architecture v4.11 physical UAT; it does not
change Aggregator, inventory, or global-snapshot authority. The documented
package contract is already GREEN through the isolated 10/10 proof. Before the
physical run, `invSys.Admin.xlam` must expose the existing Admin XLAM command
as an `ADMIN_MAINT`-gated **Aggregate Global Snapshot** action. It must use the
selected warehouse configuration's `PathSharePointRoot`, write only the
advisory/read-only global snapshot, record the administrative attempt, and
give actionable setup guidance when the publish root or source snapshots are
absent. It must not create, alter, or select a warehouse runtime root.
The access control introduces three Admin procedures (public Ribbon handler,
authorized action, and published-snapshot preflight); this is the explicit
Slice 4bd maintenance exception. Static rescanning must show no new dynamic
call or unresolved-call paths.

The physical acceptance requires two real Windows/Excel computers and **two
distinct NAS-backed warehouse runtime roots**, with distinct WarehouseId and
station identity. Two station rows inside one warehouse configuration, or a
historical local `C:\\invSys\\...` configuration, do not satisfy this proof.

Before the visible run, each station must prove the current five-package
manifest is installed, Server Sign In reaches its own runtime root, Send To
selects only its own warehouse, and invSys Sign In identifies its own station.
Each warehouse then receives/processes/publishes a small approved test receipt
under its own exact new System_Key. HQ runs aggregation only after both
published snapshots are present. The Global Inventory Snapshot must visibly
identify itself as advisory/read-only and retain each WarehouseId and each
distinct System_Key; it must never alter either source warehouse.

The final catch-up step receives/processes/publishes an additional approved
quantity at only one warehouse, reruns HQ aggregation, and proves the changed
warehouse advances while the other warehouse remains unchanged. Evidence must
record package versions, the two redacted station identifiers, the two
WarehouseIds, source snapshot timestamps/hashes, both global-snapshot values,
and an explicit local-authority unchanged check. It is local runtime evidence
only and must be redacted/ignored by Git. The proof harness must write a failed
`Harness.Exception` check for any automation fault; an incomplete report is
not GREEN.

2026-09-01 read-only NAS preflight found one NAS-backed `WH1` target with two
station entries and a separate historical `WH80` configuration rooted at
`C:\\invSys\\WH80`; therefore no second connected warehouse root is currently
available for this physical UAT. Do not repurpose either target or create a
live warehouse without the operator's explicit staging approval.

### Slice 4be -- approved implementation: shared Events and How-To/Diagnostic Action Paths

**D5 prerequisite approved 2026-09-06; implemented and technically validated.** Following
the D5 explanation, the user approved a read-only Core configuration API plus
a separate authorized headless Core write service. This replaces the earlier
Admin-owned-writer proposal. Admin remains UI/orchestration; existing scalar
Settings writes require `ADMIN_MAINT` at the command boundary, and the validated
Production UOM publication retains its narrowly scoped `PROD_POST` path. D18's
detailed Event Viewer/Action Path contract was separately approved on 2026-09-07.

Implemented `modConfigCommands`, non-mutating `modConfig` reads, Admin Settings
save routing and the validated UOM publication route. D13 recorded behavioral
RED before implementation (5 failures/10 checks), expanded baseline comparison
(9 failures/17 checks), a later required-schema RED (1 failure/18 checks), then
18/18 GREEN through packaged Admin-generated
fixtures and real Settings/UOM form handlers. All five projects compile;
live-role regression passes 48/48 and the ordered Release 1 chain passes 30/30.
The live Config fixture now checks a non-creating read before explicit setup,
replacing its superseded expectation of read-side bootstrapping.

Compile gates also exposed pre-existing Shipping diagnostic name references,
a Production parameter typo and two unqualified event constants. Repairs bind
to existing procedures/Core constants without changing identity or workflow.
Layout and Viewer pass; static candidate/duplicate/dynamic-call counts do not
grow. The actual Settings save was captured and inspected. See the maintained
[D5 evidence](../../invSys_fork/tests/integration/plan022_slice4be_d5_config_commands_results.md)
for exact commands, package hashes and final launcher/reusable Production gates.
Fresh human UAT and the separate comprehensive Event Viewer acceptance remain open.

**2026-09-07 status: detailed synthesized contract APPROVED; 4be.1 foundation and Receiving Add/Confirm candidates GREEN; comprehensive coverage pending.**
The user approved Architecture D18's shared Events, How-To/Diagnostic and Compare
both contract, dedicated Event Tracking Settings and personal view choice.
Architecture now explicitly replaces curated-only D18 and the after-R1 Viewer
deferral while retaining authoring/search/version/export/import. The defaults,
bounds and acceptance rules are approved; do not request that approval again.

**Semantic inheritance:** This plan and the control catalog may discover
controls, clarify implementation and strengthen tests within Architecture v4.11.
Record governing rule, owner, verification and rationale for each discovery.
A contradiction or material weakening requires an explicit approved decision
in Architecture before implementation; neither this plan nor a handoff can
silently authorize it. No current implementation plan switch is made:
`023 Slice 4be Critique.md` is user-supplied advisory material, not Plan 023.

**Critique disposition (considered 2026-09-07):**

| Advice | Decision under approved D18 |
|---|---|
| Events are facts; each event has an owner | Adopt as clarification. Catalog logical OwnerId and EventCode, retain owner decisions and canonical writes. Publication cannot invoke a workflow. |
| Structured payload, severity and effect | Adopt bounded fields: stable EventCode distinct from instance RecordId/ActivityId, Severity and Changed/Unchanged/Unknown DataEffect. Owner-supplied facts determine effects; unknown or partial commit never becomes rollback/success by inference. |
| Explain severe errors clearly | Adopt fixed catalog UserMessage/NextStep and sanitized cause codes; preserve existing error/confirmation handling. Reject arbitrary TechnicalDetail/raw Err.Description, paths and secrets. |
| How-To advisory; diagnostics bounded | Retain approved guide/evidence comparison. Broader runtime connectivity/lease inspection and local UI repair are not added by the critique. |
| Guide/Navigate/Inspect/Retry/Repair/Override classification | Do not adopt executable Action Path types. D18 prohibits replay/mutation; ordinary authorized operations remain with workflow owners. RetryAllowed matches an observed retry and never executes one. |
| Recursion/deduplication/cancellation | Adopt observation re-entrancy guard, idempotent same-record retry, per-ActivityId tracking-notice coalescing, and no loss of distinct user attempts. Logging cannot generate recovery commands or reopen forms. |
| Discovered safeguards feed the catalog | Adopt as required discovery metadata and protecting public-handler assertions. Further architectural changes require normative decisions, not hidden plan exceptions. |

The user-created critique file and handoff 067 edits are preserved separately;
their working-tree presence is not an implementation-plan pointer change.

**4be.1 D13 entry, 2026-09-07:** Fresh unchanged-package baseline passes the
cold-start dependency check, all five explicit compiles and 81/81 packaged
smoke. The opt-in `Test-Slice4beConfigCommands.ps1 -CheckActivityEvidence`
reuses real packaged Admin Settings and Production UOM handlers and records
**19 PASS / 12 FAIL**: all 18 D5 assertions plus the direct-service/non-user
observation check pass; three real action cases lack the new activity evidence.
The 12 failures are four assertions for each missing case, not 12 independent
root causes. See [first activity RED](../../invSys_fork/tests/integration/plan022_slice4be_activity_red_results.md).
That initial RED changed no runtime source/XLAM and claimed no GREEN or
comprehensive coverage. Initial Excel-open precondition failure was resolved before the RED
run and is excluded from behavioral evidence.

**Execution -- commit/push each completed subslice:**

4be.1 implements the D18 observation foundation first, using Config tables
`tblEventTrackingPolicies` / `tblEventTrackingControls` (both absent -> built-in
version 0; malformed/partial -> tracking unavailable). The later 4be.2 editor
will write that same approved policy. The first wire format uses GUID instance
IDs and hash-verified UTF-8 JSON as clarified in D18. Storage failure must remain
visible while ordinary authorized commands still run. Initial expanded tests
add independent hash validation and a real Settings command against a blocked
training-store path before runtime implementation.

**Foundation candidate checkpoint, 2026-09-07:** The first two catalogued
handlers now append correlated, hash-verified observations through headless
Core (code commit `03f7f20`, pushed). Supplemental boundary RED was 62 PASS / 5 FAIL (strict JSON types,
calendar-valid UTC, policy values and sequence ordinals); real-handler
stale-session RED was 68 PASS / 2 FAIL. The candidate focused suite is
70/70 GREEN, preserving all 18 D5 checks. Candidate cold-start and five-project
compile pass; packaged smoke is 81/81, live-role is 48/48, and the ordered
Release 1 chain is 30/30. Viewer/layout pass; all-role launchers are 3/3 and
reusable Production including fresh restart is 2/2. Full 4be.1 control coverage and later
subslice gates remain open. The candidate has not replaced `deploy/current`.
See [activity foundation evidence](../../invSys_fork/tests/integration/plan022_slice4be_activity_foundation_results.md)
for current regression status and boundaries. Both Settings Save Value and
Production Retrieve UOM Catalog reject stale captured sessions at the role
command boundary; tracking failure alone does not block an authorized command.
Core session state and role-owned action orchestration use bounded modules;
Auth shrinks by two lines and Production's existing form does not grow.

**Receiving coverage RED, 2026-09-07:** The unchanged foundation candidate now
has a focused packaged Receiving test: **30 PASS / 8 FAIL**, with no harness
exceptions. All 18 D5 checks and twelve independent Receiving business/binding
guards pass. The real Add/Confirm Writes handlers produce two owner-generated
events in an applied case and two queued-but-unapplied events in a controlled
pending case. Both lack the required activity attempt/result and exact source
correlation. The eight failed assertions describe these missing observations;
they are not eight independent defects. No runtime implementation or package
changed. See [Receiving activity RED](../../invSys_fork/tests/integration/plan022_slice4be_receiving_activity_red_results.md).
The next implementation must obtain source references from the Receiving owner,
preserve captured context and distinguish submission from Domain application.
Existing batch-level success is insufficient per-event diagnostic proof. D18's
approved observation semantics govern this discovered control; no architectural
contract change or further approval is implied by the test checkpoint.

The D18 Receiving wire clarification now specifies exact WarehouseId/SourceKind/
EventId/SubmissionState references. Submitted confirms queue acceptance only;
Unknown retains uncertain submission after failure. CONFIRMED/PENDING retain
Unknown Domain effect until owning application evidence exists. Catalog 2 adds
Receiving while reads preserve supported catalog-1 definitions/policies; an old
saved policy does not implicitly enable a new control. The focused extension
must reject malformed references, retain valid prior records/policies, and
exercise stale-session and tracking-store failure through the actual form.
These are semantic-inheritance clarifications implemented in the candidate.

**Receiving candidate GREEN, 2026-09-07:** The combined packaged suite passes
**133/133**, retaining all 70 foundation checks. Expanded RED was 54 PASS / 12
FAIL before implementation. Actual Receiving confirmation now observes through
a role-owned controller and calls the existing posting service once. Applied,
pending, stale-session and unavailable-store cases protect exact source
references, truthful Unknown Domain effect, binding and business behavior.

The test also discovered an implicit Config setup ownership leak during Admin
Seed. Focused ownership RED was 61 PASS / 10 FAIL; the new ownership assertion
and nine observation/read assertions failed. Core's implicit resolver now
reports newly opened Config to the existing cleanup path. Pre-existing Config
and unknown columns remain intact; policy reads preserve file bytes and refuse
unrelated dirty open Config. This restores D5 ownership discipline without
changing setup permissions or observation authority. Compile 5/5, packaged
81/81, live-role 48/48 and full-chain/restart 30/30 pass. Four actual Receiving
form captures show completed/pending work, stale rejection and tracking failure;
human acceptance remains pending. See
[Receiving candidate evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_activity_results.md)
for the expanded focused run and complete checkpoint gate results. Viewer,
Production layout/window states, three launcher checks and both reusable
Production/restart aggregate checks pass. Static candidates (1,077), duplicate
groups (192) and dynamic-call counts (45 unresolved/eight literal) are unchanged;
all 28 oversized-module ratchets hold. Comprehensive 4be.1 coverage and later
subslices remain incomplete.

**Receiving owner-outcome verification, 2026-09-07:** Unchanged candidate code
`1689bff` now passes **190/190** combined focused checks, retaining the previous
133. Actual Confirm handlers cover capability removal after staging, invalid
second-row validation, and a real saved queue with its acknowledgement withheld.
Denied has no source references/Unchanged effect; partial validation remains
Unknown; uncertain submission retains every original EventId with Unknown state.
A second explicit Confirm applies each original event once, gets a distinct
ActivityId and leaves previous observations byte-identical. Eight Receiving
captures have been inspected. No new runtime RED is manufactured for behavior
that already passes; two fixture/instrumentation failures were excluded.

Controls v1.65 maps remaining Receiving command/navigation/selection identities
as reserved and pending, plus explicit exclusions and the hidden worksheet
compatibility entry's reachability question. It corrects historical form-only
scope and receipt-identity wording to D18/current ownership; no architectural
change is proposed. Next, establish actual-handler RED for receipt Add Selected
and disposition Add/Confirm coverage before expanding the runtime catalog.

**Receiving staging/disposition D13 entry, 2026-09-07:** The expanded packaged
test is **206 PASS / 15 FAIL** on unchanged candidate code `1689bff`. All previous
190 checks pass. The new failures are missing Add Selected, Add Disposition and
Confirm Dispositions observations; independent actual staging and all three exact
RETURN/DUMP or receipt event applications pass, as do captured-workbook,
unknown-column and direct-service negative-attribution guards. The RED is preserved
in the ignored `slice4be-receiving-activity/staging-red.json` report.
An expanded pre-implementation run adds actual invalid-quantity and protected
staging-sheet cases: **210 PASS / 35 FAIL**, no harness errors. Each independently
preserves staged row values and the existing cause; missing activity must
distinguish REJECTED/Unchanged from FAILED/Unknown. See the code repository's
`tests/integration/plan022_slice4be_receiving_staging_results.md`.

Architecture D18 now explicitly records the catalog-3 staging/disposition
refinement before implementation: STAGED describes a confirmed local staging
change with empty submission references; disposition confirmation retains exact
submission evidence and Unknown Domain effect. Controls v1.66 records the same
definitions. The implementation target is Core's versioned catalog/reference
validation and Operations' actual Add/Confirm handlers; staging/business ownership
and existing package boundaries remain binding.

**Catalog-3 candidate implementation, 2026-09-07:** First focused GREEN is
245/245; expanded compatibility, stale Add, optional store-failure and readable
catalog-3 tests pass **262/262**. Receipt and disposition confirmation each retain
all four independently applied exact events, including unobserved/direct staging
without invented user actions. Eleven new actual form captures were inspected.
Five explicit compiles and 81/81 packaged smoke pass. Static candidates (1,077),
duplicate groups (192), unresolved dynamic calls (45) and literal calls (eight)
do not regress; all 28 oversized ratchets hold and Receiving shrinks 24 lines.
Controls v1.67 is synchronized. Live-role 48/48, full-chain/restart 30/30,
Viewer, Production layout/window states and all three packaged launchers pass
on the new candidate hashes. Two initial reusable Production runs and an
independently built/compiled pre-change baseline fail with native Excel/RPC
interruption. After cooperative dialog-observer shutdown replaces forced job
termination, the unchanged candidate passes two independent full reusable
Production/clean-restart runs, each 2/2, and all three launchers again. Both full
runs close Excel normally, preserve all five package hashes and record no Excel
native crash event. All checkpoint gates pass. This supports the harness
hypothesis without establishing the native crash's cause; failed runs remain in
the evidence. Comprehensive coverage, later 4be subslices and human acceptance
remain incomplete; no accepted deployment or NAS rollout is implied.

**Receiving Refresh/Clear D13 entry, 2026-09-07:** The next 4be.1 checkpoint
retains the 262 GREEN cases and tests both Receiving and Returns through their
actual Refresh/Clear handlers. Expected RED is missing catalog-4 observations,
stale-context owner calls and false refresh success after an owner returns False.
Generated fixtures additionally protect local-only writes, unknown columns,
unchanged authority bytes, internal/direct-call negative attribution, optional
tracking failure, empty Clear and a real protected-second-table partial failure.
The first run is 322 PASS / 74 FAIL, with two combined fixture-header checks
excluded from intended RED. Expanded current-table header, closed-workbook and
catalog-3 policy checks produce **365 PASS / 74 FAIL** with no harness exception;
all earlier 262 checks remain GREEN. Header and unrelated-workbook guards now
pass, and all remaining failures match the intended missing behavior. The initial
catalog-4 candidate subsequently passes **439/439** focused checks, five builds,
explicit compiles and cold start. It is not a completed checkpoint: source
inspection found that the read-model owner can return True while retaining
cached inventory or loading a stale fallback. Additional actual-handler tests
now exercise both source states on Receiving and Returns before a freshness fix.
Architecture D18 clarifies owner-returned REFRESHED/STALE/FAILED primitives;
STALE is Warning/Changed for local metadata/projection work, with empty source
references and preserved visible cause. Boolean compatibility and existing
snapshot selection remain unchanged. This is a discovered outcome refinement
under approved semantic inheritance, not a new diagnostic authority. The Core
read-model owner and its declared bridge join the bounded implementation scope.
The expanded test is **463 PASS / 28 FAIL**, followed by **491/491 GREEN** on
`deploy/validation-receiving-freshness`, retaining all prior checks. All 22
Receiving Refresh/Clear/freshness captures were inspected. Five builds/compiles,
cold start, packaged 86/86, live-role 48/48, full chain/restart 30/30, Viewer,
Production layout and three launchers pass. Dedicated reusable Production is
2/2 GREEN, including full reusable/Chai cases and clean-process restart. All five
candidate hashes remain unchanged; Excel is closed. The three Production layout
captures were inspected. Static candidate/duplicate/dynamic counts and
all 28 oversized-module limits hold; the read-model owner shrinks three lines.

The initial packaged gate's 81 checks passed but it left a Designs save prompt.
Its validator lacked runtime-root isolation and complete workbook cleanup.
The corrected harness retains the 81 checks and adds five explicit Config
fixture/root/cleanup checks; both Excel sessions close and the saved default
Designs file remains byte-for-byte unchanged across the corrected runs. An
intermediate 84/85 run exposed the old external Config fixture dependency; it is
harness evidence, not D13 product RED. No runtime contract changed for this
correction. Comprehensive coverage and user comparison remain incomplete. See the code repository's
`tests/integration/plan022_slice4be_receiving_local_results.md`.
Architecture D18 records the precise discovered-control outcomes
before implementation; controls v1.70 records the candidate evidence. Core
catalog/outcomes, read-model owner/bridge, the Receiving form/controller and
existing local Clear owner implement this checkpoint. No new business permission
or canonical writer was added.

**Receiving Open/Close D13 entry, 2026-09-07:** Continue 4be.1 from the preserved
491-check catalog-4 candidate. Architecture D18's Open/Close clarification
governs catalog 5: owner-confirmed OPENED/REUSED/FAILED and explicit CLOSED,
empty business references, no internal-unload attribution, and session/workbook
validity before reuse. Add actual generated Ribbon callback and Close/QueryClose
handler tests before runtime changes. Expected RED is missing observations and
reuse of an old-session form. Preserve earlier GREEN checks, optional-store
failure behavior, direct-call exclusion and unrelated workbook/authority bytes.
The Core catalog/outcome definitions, Operations launcher/form and generated
Receiving Ribbon route are the bounded expected runtime scope. First packaged
RED is 515 PASS / 37 FAIL, followed by **531 PASS / 47 FAIL** across 578 checks
with no harness exception and all prior 491 GREEN checks preserved. The failures
are absent observation pairs, stale-session reuse and missing tracking-failure
notices. Valid catalog-4 policies, both dismissals under store failure, staged
identities/unknown values, authority bytes and same-name workbook reopening
pass. Reopening is a regression guard, not another confirmed defect. Three
lifecycle captures are inspected; static metrics and all 28 existing size limits
hold. These RED packages remain preserved. Catalog-5 implementation then passes
**578/578** on an independently built candidate, including five compiles and cold
start. Ribbon dispatch explicitly marks the user action; the existing launcher
owns provisioning/binding and checks captured session/workbook before reuse.
Close completion follows dismissal, with one pair and retained tracking notices.
After GREEN, pure header-formatting and message helpers are moved within Receiving
to keep existing module limits. Source regressions pass. A rebuilt candidate
adds expanded native-window close checks before the remaining release gates.
The initial caption-lookup guard yields 542 PASS / 1 harness failure. Exact
captured-form handle lookup then exposes product RED: **592 PASS / 4 FAIL**, with
all 578 prior checks retained. Cancelling native close while invoking nested
Unload leaves the form open but reports CLOSED. D18's committed-dismissal rule
governs the first correction: allow the native close, release the cached form
reference, and finish its already-requested observation on termination. Internal
unloads start no activity. That candidate compiles, but its first narrow
run stops at **131 PASS / 1 harness exception**: the native window disappears,
then Excel loses COM connectivity and Windows records a VBE7 access violation.
Unchanged catalog 4 completes 129/55 without crashing. Suppressing termination
publication still crashes (125/7); retaining the reference avoids the crash but
delays completion and its notice (176/8). These are diagnostic runs, not GREEN.
D18's UI-dismissal clarification governs the next correction: synchronously hide
on explicit native close, finish that dismissal, invalidate reuse without
releasing the cache during QueryClose, and allow native teardown. The button
retains direct Unload completion. The new five-package candidate passes cold
start, all five compiles and **596/596 GREEN**, preserving every 491/578 prior
check by name. Both native closures and immediate evidence pass without a crash;
three operator captures are inspected. Static metrics and all 28 size limits
hold. Packaged validation passes 86/86, live roles 48/48, full chain/restart
30/30, Viewer and layout pass, and public launchers pass 3/3. The full reusable
Production run stops at released Process edit/export with one RPC harness failure
and an ntdll fault. The unchanged catalog-4 comparison passes 2/2; a candidate
retry fails at a different stage with the same native signature. Neither is
proven product RED. A scoped harness test then exposes automatic acceptance of
an OK/Cancel confirmation (6/1). A native modal-only observer passes 7/7, keeps
ordinary controls/other processes untouched and retains cooperative shutdown;
public launchers pass 3/3 with it. The candidate's full Production and clean
restart then pass 2/2 with unchanged XLAMs/assertions and no new Excel crash.
Final candidate and catalog-4 hashes remain unchanged; Excel is closed. The
crash's cause remains unproven and failed runs remain recorded. The technical
Open/Close checkpoint is complete; comprehensive coverage and human UAT remain
open. Controls v1.76 and code evidence
`tests/integration/plan022_slice4be_receiving_lifecycle_results.md` record this
checkpoint. The prior Refresh/Clear checkpoint remains verified.

**Receiving navigation/selection D13 entry, 2026-09-08:** Continue 4be.1 from
the preserved catalog-5 lifecycle candidate and all 596 GREEN checks. D18's
navigation refinement and controls catalog reserve thirteen page/list/choice
identities, fixed captions, Navigation/default-off policy, UI-only SELECTED
outcomes and empty source references. This is semantic inheritance of approved
coverage/input-exclusion rules, not a new business contract. Before runtime
changes, extend the packaged harness to distinguish deliberate mouse/keyboard
selection from programmatic changes, initialization/refresh and dependent fills.
Expected RED is absent registered/correlated optional observations. Native input
must independently reach its actual control; unavailable input or a broken
fixture is harness failure, not product RED. Protect stale/captured context,
older/disabled/malformed policy, tracking failure, exact keys/unknown columns and
business authority bytes. Recording integration remains separately pending.

**Navigation focused GREEN, 2026-09-08:** Full unchanged catalog-5 RED is 632/74
with all 596 prior checks GREEN; expanded mouse/keyboard RED is 118/122 in a
separate diagnostic run. Catalog-6 candidate `validation-receiving-navigation`
builds/compiles all five packages and passes 253/253 focused checks. The existing
form handlers use a typed navigation boundary and form-local input state;
dropdown open/close clears pending input, including same-choice/cancel paths.
Full regressions and release gates remain pending. See
[navigation evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_navigation_results.md).

**Discovered D14 identity blocker, 2026-09-08:** Navigation's full form suite is
767/767 and its additional closed-workbook focused run is 257/257, but the ordered
Release 1 chain is 29/30: packaged Admin Generate succeeds and Seed is rejected
with DUPLICATE_SYSTEM_KEY. Preserve that failed chain. The existing Core fallback
reseeds VBA Rnd for each generated identity; investigate through the unchanged
packaged CreateSystemKey boundary before altering implementation. The focused
`Test-PackagedSystemKeyGeneration.ps1` checks 50,000 creations both normally and
with caller-side RNG resets, reports counts only, and preserves package bytes.
The existing actual Admin Generate/Seed chain remains the operator-path gate;
the identity probe supplements it. Expected RED is duplicate keys. This restores
D14's existing global uniqueness requirement, without changing existing identity
values, migrating data or relaxing Domain duplicate rejection. Keep 4be.1 active
and incomplete; no navigation technical completion until this blocker and all
remaining release gates are GREEN.

The packaged identity probe is meaningful RED: 50,000 normal calls yield 35,017
unique keys and 14,983 duplicates; the ambient-RNG reset control passes and
package bytes are unchanged. After that RED, Core extracts its existing native
GUID implementation into modSystemIdentity. Role entity/event creation and the
archive collision suffix call it directly; Training keeps its representation and
failure wording through its existing wrapper. Reviewed callers of the removed
fallback/normalizer are accounted for. Candidate
`validation-receiving-navigation-identity` builds/compiles all five projects and
passes 3/3 focused checks (100,000 creations, zero duplicates/blanks). No stored
identity changes, migration or relaxed Domain rejection are introduced. The
source Create Warehouse harness initially omits the new dependency; correct all
sixteen explicit import lists (parser PASS), then restart the chain. Its source
integration now passes 15/15. This compile dependency failure is not product RED.
Final static evidence is 175 components/5,500 procedures with unchanged candidate,
duplicate and dynamic-call counts from the navigation review below. All 28
pre-existing module caps still hold.

**Final navigation/identity technical checkpoint GREEN, 2026-09-08:** The corrected
replacement chain completes 30/30, including Admin Seed and clean restart
reconciliation. The full activity/form suite is 771/771, retaining all 767 prior
check identities and incorporating the four closed-workbook guards in one run.
Five-package build/compile/cold start, smoke 86/86, live roles 48/48, Viewer,
three-size/five-page layout/native transitions, public launchers 3/3 and full
ProductionReusable/clean restart 2/2 all pass. No reduced Production flags are
used. Receiving and Production captures are inspected; duplicate image display
was checked individually and by file hash before accepting visible evidence.
Replacement/first-navigation/catalog-5 hashes remain unchanged, as do verified
catalog-4 hashes. Excel is closed with no Application Error in the final
Production interval. Code **77a0851**, controls v1.78 and the linked evidence
record the completed technical checkpoint. Comprehensive 4be.1 coverage and
human UAT remain open.

**Explicit maintenance exception for this navigation checkpoint:** Retain the
seventeen required native event roots and three new normalized duplicate groups
consisting only of short input-reset/mouse-input/form-dispatch callbacks. These
distinct MSForms interfaces share one typed controller and contain no duplicated
business logic. Raw duplicate groups remain reported as 192 -> 195; total
candidates 1,077 -> 1,097 include retained roots. Existing 28 size limits and
dynamic-call limits remain binding. Correct the missing DropButtonClick class
event classification using the native evidence; do not delete its live handler.
This explains functional adapter growth under D18's no-unexplained-growth rule
without changing runtime/architectural authority or accepting comprehensive 4be.

**Receiving launcher-denial D13 entry, 2026-09-08:** Continue 4be.1 under D18's
launcher-denial clarification, preserving the 771-check navigation/identity
candidate and 105-check surface discovery. The protecting packaged test is
`Slice4beReceivingLauncherDenial.ps1`, entered by the actual generated Operations
Ribbon callback. Expected RED is absent correlated REQUESTED/DENIED observations
and optional-tracking notice on denied launch. Guard rejection, existing auth
decision, no provisioning/form opening, untouched staging/identities/unknown
columns, direct-guard/polling and pre-sign-in exclusions must already hold.

The intended change keeps the actual Core cached RECEIVE_POST guard, moves its
single action-time invocation into the typed Receiving entry before any owner
work, and leaves generated getEnabled capability mapping intact. Generator
metadata must explicitly identify this already-guarded entry; other controls
retain their generated guard. Register RECEIVE_OPEN_DENIED for the existing
catalog-6 control with Blocked/Unchanged and empty references. Capture context
before the guard, preserve denials during optional-store failure and never
rebind an interrupted completion. No new permission, business owner or canonical
schema is authorized. Architecture, Plan and controls must be synchronized before
runtime implementation. Full packaged and technical gates remain required after
focused GREEN; native worksheet input remains independently pending.

The unchanged-package focused runs are **108 PASS / 19 FAIL**, then expanded
**112 PASS / 21 FAIL**, with no harness exception. The expanded RED includes
captured-session interruption. Architecture/Plan/controls clarification was
pushed as docs **f918687** before implementation. The separate denial candidate
passes focused **133/133**, native-dialog **137/137** and full activity
**845/845**, retaining all earlier 771 and surface-discovery 105 check identities.
Both actual fixed denial/tracking dialogs were captured and inspected.
Existing authorization, no launch owner entry, identity/unknown-column and
workbook preservation, exclusions and optional-store behavior pass.

Five packages build/compile/cold-start; smoke **86/86**, live-role **48/48**,
ordered Release 1 chain **30/30**, Viewer, three-size Production layout and
public launchers **3/3** pass. Current static ratchets preserve all 28 prior
module caps and component/candidate/duplicate/dynamic-call counts. The historical
Slice 12 validator retains the same six failures against both old/new baselines;
it is not reported GREEN. The full reusable Production gate remains unresolved:
two candidate runs crashed in native Excel at different stages, while the
preserved prior candidate passes **2/2** in the same harness. Core-only and
Operations-only substitutions into the other prior packages each pass **2/2**;
these are diagnostic combinations, not accepted release sets. A clean rebuild
passes cold-start/five compiles and matches all 168 extracted VBA components,
but also crashes at batch scale. An unsaved source-preserving Operations
recompilation probe likewise crashes at batch scale. Both report the same native
fault. No cause, regression exemption or checkpoint completion is established.
The subsequent provenance comparisons pass 2/2 for candidate Core+Operations
with prior Domains, all current executing packages relocated, and the original
candidate at its original location. Actual initial/restart dependencies resolve
within each intended directory. These inspection variants do not prove a repair;
Admin is present but not loaded by this harness. A one-second pause without
project inspection still fails at batch scale. Saving resolved copied packages
also fails there despite unchanged source and five successful compile checks.

Corrected bounded native observers capture a preceding `0xc0000005`, then
`0xc0000028` through VBE7/ntdll; both full runs fail at the variable-quantity
boundary. The first observer run is inconclusive because it could swallow other
exceptions; stronger synthetic calibration now passes 15/15. An unsaved fixed-stage
probe fails at ReleaseSource in the released Process edit/export test. None of
these traces identifies the faulty VBA statement or supplies behavioral RED.

A fresh copied set with all five projects explicitly recompiled **before** saving
preserves all 168 source components and passes its initial full workflow, but
fails clean restart: 1 PASS / 1 failure. It is not an accepted build fix or release
set. A full rerun with 20 fixed harness markers around restart teardown/loading/
actions passes 2/2 without project inspection or VBA mutation. Timing/intermittency
remains unresolved. Subsequent exit verification finds the outgoing process
terminal in three focused teardown cases and a full workload/restart run (2/2),
with a distinct restart identity and only in-memory markers. This provides no
support for an exit-wait repair. The original candidate still fails at batch
scale under a calibrated native observer: the preceding c0000005 is a first-chance
execute violation at an unresolved site. Metadata calibration is 24/24; no raw
addresses, memory contents, stack arguments or operational values are persisted.

A separate apparent RibbonX build-order concern was disproved by live calibration:
SaveAs to XLAM writes a copy while the source workbook retains its original name
and empty path. The actual saved artifact is not loaded. The unchanged builder
passes 7/7 external-edit boundary checks; its exact detector passes 5/5 calibration,
including rejection of an explicitly opened XLAM and byte preservation. The first
probe's enumeration/JSON assumptions were harness errors, not behavioral RED.
No build reorder, new exit wait, runtime change or architectural amendment was
implemented. The new boundary test protects already-correct package assembly.
The memory-metadata extension subsequently passes 27/27 observer checks after
24 PASS / 3 FAIL for absent labels, plus 10/10 live-allocation checks. It records
finite state/type/protection labels without reading page contents or persisting
addresses. Original-candidate initial observation and compiled-candidate restart
observation each complete the full workflow 2/2 without capturing an exception;
the first observer's 180-second window covers only part of the full workflow.
A bounded original-candidate capture repeat also passes 2/2 without an exception.
All three runs end with Excel closed and all 20 package pins preserved.
The failing target's memory state remains unknown. Finer fixed markers are then
calibrated around the existing Process Release queue/processor/status/list-refresh
boundaries and combined with native observation in an isolated unsaved diagnostic.
No runtime/build repair is established.
See [memory diagnostic evidence](../../invSys_fork/tests/integration/plan022_slice4be_production_memory_diagnostics.md).
These are diagnostic results, not permission to weaken or replace the full gate. See the
[exit/execute evidence](../../invSys_fork/tests/integration/plan022_slice4be_production_exit_and_execute_results.md).
All 20 pinned packages
across the four preserved sets remain unchanged; Excel is closed at this check.
No runtime/architectural contract changed in this diagnostic work. D12/D13/D18
and every remaining release gate remain binding. See the continued
[diagnostic evidence](../../invSys_fork/tests/integration/plan022_slice4be_production_provenance_diagnostics.md)
and the original code evidence
`tests/integration/plan022_slice4be_receiving_launcher_denial_results.md`.

**Recovery/current standard gate, 2026-09-12:** After the user's reported outage,
the old diagnostic handle is absent and its retained report proves a setup failure
before callbacks. Excel's `.Name` -> `.name` normalization defeated a case-sensitive
marker match. Actual-loaded-source calibration goes 0/3 -> 3/3; placement remains
13/13 and error-preserving transport 5/5. The corrected full diagnostic passes 2/2
and records 11 complete Release sequences, without a native fault capture.

Input-desktop access and a foreground window are now available. In this changed
environment the unchanged standard validator against the original pinned candidate
passes **2/2**, including clean restart, without debugger, markers, project
inspection or reduced-workflow flags. This is a current standard-gate GREEN;
the earlier native failure remains unexplained and no causal fix is claimed.
All 20 pins remain unchanged; Excel is closed. Native 27/27 and memory 10/10
calibrations were rechecked after recovery. No runtime, builder, static baseline
or architecture changed. The subsequent native worksheet caller checkpoint below
advances coverage discovery; comprehensive 4be coverage and human acceptance remain
incomplete. See
[Release/recovery evidence](../../invSys_fork/tests/integration/plan022_slice4be_production_release_boundary_diagnostics.md).

**Receiving worksheet surface discovery, 2026-09-08:** D18's existing coverage
and exclusion-evidence rule governs this test-only discovery. The preserved
navigation/identity candidate passes **105/105** through the real Receiving
Ribbon callback: provisioned/reused support sheets remain VeryHidden, but an
eligible saved workbook with ReceivedTally as its only visible sheet retains
the visible Confirm Writes button and expected handler through save/reopen.
Captured workbook, unknown header, saved bytes, authority files and unrelated
workbook checks pass; all five package hashes are unchanged. This is not a new
runtime contract, activity implementation GREEN or a repeat of all release gates.

Excel normalized the assigned macro to unqualified
`modTS_Received.ConfirmWrites`; the initial four qualifier assertions were an
invalid test assumption, not a routing defect. Native-message attempts lacked
calibrated input delivery, and minimal transport tests could not establish
foreground ownership / Windows cursor positioning in the session. Neither those
failures nor the native calibration exception is D13 behavioral RED. Native
worksheet caller/handler evidence remains pending before its explicit activity
identity and owner-outcome implementation. Do not exclude the control solely
because the usual launcher state hides its sheet. Architecture D18 is unchanged;
controls v1.79 records the discovered surface without weakening any rule. See
[surface evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_surface_results.md).

**Receiving native worksheet caller proof, 2026-09-12:** Minimal disposable input
calibration is **8/8**. The original pinned launcher-denial candidate then passes
**115/115**, retaining all prior 105 surface check identities. Native clicks on
the visible sole-staging-sheet and saved/reopened controls enter the actual
Operations `ConfirmWrites` handler with exact `btnConfirmWrites` caller identity.
Saved bytes, captured form binding, unknown header, authority and unrelated
workbook checks pass. Both worksheet captures were inspected. Earlier coordinate,
window/context and COM setup failures remain harness failures, not product RED.
The working harness explicitly activates the disposable workbook's own window;
its Visible property was already true before preparation. All 20 candidate pins
remain unchanged, Excel is closed, and no runtime/architecture/static baseline
changed. No human UAT or repeated full release gates are claimed. Under D18's
existing coverage rule, next refine the worksheet control identity and protect
missing activity through meaningful native-action RED before implementation.
Controls v1.86 and [surface evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_surface_results.md)
record this checkpoint; the broader 4be.1-4be.6 scope remains open.

**Receiving worksheet activity D13 entry, 2026-09-12:** Continue 4be.1 under
the approved D18 coverage/owner-fact rules, following native115/115 and minimal
input8/8. The normative worksheet clarification registers catalog7 control
RECEIVING_WORKSHEET_CONFIRM and its fixed worksheet surface/prefix, preserving
catalogs1-6 and existing business authority. Actual native entry must capture its
worksheet/workbook and trusted session, retain that binding, and recheck context
before posting. Direct macro/service calls remain unattributed. Owner rejection
retains Unknown effect, since validation can follow local staging normalization.

First focused RED: actual worksheet clicks in the existing sole-visible-sheet
and saved/reopened cases lack attempt/result records, stable correlation and
registered-readable metadata; the same programmatic macro call must add no user
activity. Extend coverage to exact multi-event submissions, pending/uncertain
results, denials, changed context, unavailable tracking, older policies and
unknown-column/authority preservation before claiming worksheet implementation
complete. Expected runtime changes are bounded to the Operations Receiving entry/
activity controller and Core catalog/outcome/reference validation. No runtime
change or RED/GREEN result is claimed by this contract/test entry; the full
4be.1-4be.6 and Release1 gates remain required. Controls v1.87 carries the same
reserved identity and pending acceptance state.

**Receiving worksheet candidate checkpoint, 2026-09-12:** The unchanged original
package gives meaningful native activity RED at 121 PASS/12 FAIL with no harness
exception; the isolated catalog7 candidate passes the same 133 checks. Both actual
worksheet entries/callers and owner rejection are proven; direct macro invocation
remains unattributed. All five candidate packages build and compile, Operations
cold start passes, and the preserved full activity regression passes 845/845.
Source/tooling checks pass 6/6 control surface,14/14 cutover,10/10 Receiving and
62/62 tool contracts. Expanded native submission/failure/context/policy evidence
and remaining candidate release/visible gates are still required. A fixture
window-index exception during expanded testing is not behavioral RED and does not
justify a runtime workaround. The maintained [worksheet evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_worksheet_activity_results.md)
records exact local reports and pending gates. This checkpoint does not complete
worksheet acceptance, comprehensive 4be.1 coverage, or 4be.2-4be.6.

Expanded native submission proof now finishes RED175PASS/35FAIL -> GREEN210/210,
with the same210 check identities and no duplicates or harness exceptions. Applied,
Pending, UnknownSubmission, StoreFailure and OlderPolicy all enter the actual
worksheet handler; independent inbox/Domain checks retain every exact identity and
uncertainty. Unknown headers, pending custom values, Config bytes and unrelated
workbooks are preserved. Tracking failures do not block business success and their
notices reach the existing notification surface. Packaged smoke86/86 and final
static generation pass; no dynamic-call/duplicate metric or oversized-module cap
regresses. Native context/denial edge cases and remaining release/visible gates
remain open;210/210 is not full worksheet or Slice4be acceptance.

**Worksheet candidate regression status, 2026-09-12:** Live roles48/48, Viewer,
layout geometry/native-window checks and isolated full Production/restart2/2 pass.
The first full-chain run stopped after4 Admin/Seed checks with an Excel RPC/native
failure; the pinned original comparison and unchanged clean candidate rerun both
pass30/30. Native exceptions remain unexplained, including a Windows crash record
during the successful chain's time window. The candidate combined public-launcher
gate fails twice after Receiving at Production's batch-scale test; the pinned
original passes3/3. This is an unresolved release-regression gate, not full
acceptance and not evidence for a speculative source repair. A separate identical-
source five-package rebuild passes build/compile/cold-start and that same launcher
gate3/3 and focused worksheet210/210. Its remaining package regressions are being
verified separately; all30 package pins remain unchanged. No build-defect or
native-crash repair is inferred. Source/test/static/
evidence checkpoint **01891bb** is committed and pushed; it is not Slice4be
acceptance or deployment. The layout-validation screenshot also shows a
partly clipped Committed / Used heading; recheck the normal launched Production
form before assigning a layout defect or claiming visible acceptance. All outcomes
and remaining context/denial guards are retained in the worksheet evidence.

**Worksheet guard-test entry, 2026-09-12:** Continue4be.1 under the existing D18
captured-entry/owner-fact rules. Supplemental packaged native tests cover Denied,
SignedOut, SwitchWorkbook, SignOutDuringTracking, CloseDuringTracking and a change
to a different generated, signed-in warehouse during tracking. Unsaved
seams are installed before fixture/form creation; they observe existing posting-
owner counts/binding and interrupt only the optional-tracking return in isolated
fixtures. Actual `btnConfirmWrites` input remains mandatory. Protect no owner
retry or workbook redirection, no business mutation on rejected actions, exact
source references on authorized completion, preserved unknown fields, and no
invented conclusion after context loss. A failed interruption or fixture setup
is not behavioral RED. These tests add coverage without changing runtime; any
concrete failure must protect a subsequent correction before that correction is
implemented. The normal combined launcher and complete Release1 gates remain
required independently of these instrumented guard proofs.

**Worksheet resumed verification, 2026-09-12:** The prior guard process is terminal
at215/215 GREEN (the five-case predecessor was201/201). The rebuilt full activity
suite is845/845 with every prior identity retained and no duplicates; smoke86/86
also passes. The rebuilt live-role gate stops at39PASS/1harness exception during
Production Complete Run, with RPC0x800706BE and Windows native c0000028/ntdll.dll.
Recovery files were preserved through the inspected keep-for-later prompt. No
runtime repair is inferred. The pinned original comparison and unchanged clean
rebuild subsequently both pass48/48; the original failure remains recorded.
The rebuilt full chain passes30/30, Viewer passes, and layout geometry/window
checks pass. A Windows c0000005/combase.dll record during the successful chain
still prevents any claim that native stability is resolved. The rebuilt default
layout capture retains the clipped Committed / Used heading; normal populated-
form and human review remain open. Full reusable Production/restart passes2/2
without reduced flags, debugger or diagnostic mutations.
The strengthened guard passes215/215 with the actual second warehouse ID,
runtime root and signed-in actor verified, preserving all check identities. Its
first attempt stopped94/1 before native input because the fixture could not obtain
the foreground window; that setup failure remains separate from behavioral RED.
The unchanged retry required no input-helper/runtime repair. All30 package pins
and four runtime source pins match; Excel is closed. Static JSON contracts and
all28 module limits pass. Native stability, visible/human acceptance and
comprehensive4be.1-4be.6 remain open.
Code checkpoint **6c9d2a1** records the guarded215 test, rebuilt package gate
evidence and Shipping/Boxing source discovery. Runtime remains the01891bb
candidate; accepted deployment/NAS and unrelated user files are unchanged.

**Shipping/Boxing discovery, 2026-09-12:** The maintained
[source coverage map](../../invSys_fork/tests/integration/plan022_slice4be_shipping_coverage.md)
identifies current form commands, owner boundaries, navigation and unresolved
worksheet/Ribbon reachability under D18. Shared activity remains unimplemented.
The source distinguishes SHIP_POST save/make actions from ADMIN_MAINT box
maintenance; the observation layer must preserve both. This is non-contract
source discovery, not D13 RED, a new catalog version or Shipping acceptance.

**Shipping activity D13 test entry, 2026-09-12:** Continue4be.1 through the
public Shipping launcher and actual Add/Update/Remove/Send Hold/Return/To
Shipments/Shipments Sent form handlers. `Test-Slice4beConfigCommands.ps1
-CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity` adds a
separate ignored report directory and generated warehouse with a genuinely
authorized SHIP_POST user. Unsaved facades expose the current handlers without
auth bypass; only existing report presentation is intercepted. First require
observable staging results, exact selected System_Key, captured workbook,
unknown headers, Config bytes and unrelated workbook preservation. Expected
behavioral RED is missing attempt/result, stable correlation, owner/context and
sanitized activity. These initial assertions do not yet establish exact source
references or Domain application; strengthen that owner evidence and record
precise catalog/outcome definitions under D18 before runtime implementation.
No Shipping/Boxing runtime or catalog version changes in this test entry.

**Shipping activity RED, 2026-09-12:** Initial68/1 stopped before Add because
seed inventory alone did not supply a saved shippable box. The fixture now uses
real Box Designer/Box Maker handlers; an owned single-OK observer makes the test
self-contained without accepting confirmations or retaining dialog text.
The reproducible105PASS/32FAIL run expands to125PASS/40FAIL with every failure
confined to missing Shipping activity. Independent staging/key/binding/header,
Update/Stage unknown-value, Config/unrelated-book and source-read non-mutation
checks pass. Domain evidence proves one quantity2 shipment and ten-made/eight-left
balance for the exact box key. No harness exception or duplicate check exists.
All30 package/four runtime source pins remain unchanged; no runtime implementation
is claimed. [Shipping activity evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_activity_results.md)
records exact runs and the remaining negative/pending cases and contract refinement.
Code **7420b7f** is committed and pushed. Static contracts pass with unchanged
runtime175 components/5503 procedures,8 literal/45 unresolved dynamic calls and
all28 module limits; only test reachability metadata grows. Controls v1.90 records
the same RED scope. Next protect rejected/stale-context/pending/failure outcomes,
then refine the Shipping catalog and owning result envelope under approved D18
before implementing observation. Receiving and the full Release1 goal remain open
at their existing acceptance boundaries; no narrower completion is claimed.

**Shipping rejection/context D13 entry, 2026-09-12:** Continue 4be.1 with the
same packaged launcher/form fixture. Append zero-quantity Add and same-user
sign-out/re-sign-in cases after all eight normal actions. D18 already requires
pre-owner captured-session validation independently of optional activity; a new
sign-in must not revive an existing form. Unsaved test facades expose input/status
and read the actual Core session version; they neither bypass auth nor change
owning workflow behavior. Runtime and Architecture remain unchanged.

The first expanded run is 137 PASS / 47 FAIL, preserving all 165 prior check identities
and all 125 prior GREEN checks. Zero quantity reaches existing validation with
staging/canonical data unchanged, but six activity assertions fail. The stale
form changes local staging after verified same-user reauthentication: one actual
D18 contract failure beyond missing activity. No canonical event is applied;
that alone is not proof of no submission. Extend the context matrix before implementing the guard and
registered activity outcomes. See the maintained [Shipping evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_activity_results.md).
This is discovery of nonconforming runtime under the existing approved rule,
not a new permission or an architecture change. Slice4be and Release1 remain open.

The strengthened packaged run finishes 139 PASS / 49 FAIL. Calibrated unsaved counters
prove zero quantity stops before submission, but the stale form reaches both
`ShipmentsFormCommitLine` and `QueueShippingPayloadEventServerFirst`. Three
stale-session guard assertions now fail alongside 46 missing-activity checks.
Every earlier check and GREEN result is retained, with no duplicate or harness error.
Submission entry is not successful acceptance; canonical non-application does
not establish an empty queue. The required fix remains pre-owner context rejection
under existing D18. Controls v1.91 records the same scope; no new architecture
decision or runtime implementation is asserted.
All 30 package/four source pins match; Excel is closed. Static JSON contracts pass,
and runtime metrics and all 28 module limits are unchanged.
Code checkpoint **8e64268** is committed and pushed; the next test-first work is
the remaining Shipping mutation/context matrix before the D18 guard repair.

**Shipping mutation/context matrix and repair entry, 2026-09-12:** Architecture
D18's Shipping mutation-context clarification constrains the repair to seven
actual mutation controls and explicit recovery through the authorized launcher.
The new `Slice4beShippingContext.ps1` probe calibrates each healthy handler at its
existing owner boundary, then stops the owner there to test stale dispatch without
business effects. The preserved normal sequence still runs real owners and proves
staging/source/application behavior. Probe success is not Domain acceptance.

Signed-out, reauthenticated and changed-target handlers must reject visibly before
owner entry. A valid launcher must reuse its form; an explicit launch after stale
context must replace the old form on the same workbook and preserve active/held
staging. Any UI yield requires a fresh context check before owner entry. Automatic
sync must stop for the rejected form. Existing Core capability ownership and normal
Close remain binding. Focused RED precedes runtime changes to `frmShipmentsTally`,
the Shipping launcher and bounded context/timing helpers. Timing extraction, if
needed to keep the form within its size limit, retains the existing status text
and has packaged normal-action status assertions. No new activity catalog is added.

Matrix RED on the unchanged packages is **242 PASS / 91 FAIL**: all seven healthy
owner calibrations pass; 21 stale-context cases each fail pre-owner rejection and
the context notice. The prior 49 failures remain. Every prior check and GREEN
result is retained. The initial matrix stopped at a test-only availability threshold
of five for a two-unit Add (136 PASS / 50 FAIL including one harness failure);
correcting that threshold yields the meaningful calibrated RED. Explicit launcher
recovery coverage is being established separately before implementation.

**Shipping context candidate, 2026-09-12:** The explicit-relaunch extension is
246 PASS / 92 FAIL on the unchanged packages. After normative clarification
**089ab0d**, the candidate passes all46 formerly failing guard/relaunch assertions:
292 PASS / 46 FAIL overall, with only missing Shipping activity remaining and
every prior GREEN retained. The full suite still exits1; this is focused context
GREEN, not activity acceptance. [Context evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_context_results.md)
records exact tests, source changes and remaining gates. Both large Shipping
modules shrink; the bounded context/lifetime module and timing class add73 net
runtime lines. All28 size limits hold;8 literal/45 unresolved dynamic calls and195
duplicate-body candidates are unchanged. Five-package build, all five compiles,
Operations cold-start, packaged smoke86/86, live-role48/48, ordered full-chain30/30,
Viewer, combined launchers3/3 and Shipping layout1/1 pass. Full-chain business
success coexists with an Excel c0000005/combase.dll fault during its time window;
native stability remains unresolved. Full reusable Production/restart passes2/2;
the visible rerun retains292/46 with no other failure or lost GREEN. The inspected
session notice is legible; crowding near the main Shipping key editor/Add button
and adjacent headers requires baseline comparison before human layout acceptance.
All tests are terminal and Excel is closed;35 package/eight source pins match.
These results apply to the same isolated
`deploy/validation-shipping-context-guard` candidate. Accepted deployment/NAS are
untouched. Capability-loss, during-yield/closed-workbook/automatic-sync probes and comprehensive
activity/Settings/Viewer/Action Path acceptance remain open.
Code checkpoint **84a023a** is committed and pushed. This is the tested Shipping
context candidate, not Slice4be completion or native/human acceptance. Continue
the remaining context/owner-outcome tests and activity contract implementation
without dropping the preserved GREEN assertions.

**Shipping interruption/timer D13 entry, 2026-09-12:** The unchanged pinned
context candidate records **319 PASS / 49 FAIL**. All18 checks interrupting Add,
To Shipments and Shipments Sent at their real pending-status UI yield pass.
Healthy timer dispatch reaches its calibrated owner and reschedules pending work;
after actual sign-out, the registered timer incorrectly reaches that owner,
reschedules and omits the context notice (three new behavioral failures).
The other46 failures remain missing Shipping activity. Stopped-owner probes prove
dispatch boundaries, not completed synchronization or business effects.

D18's Shipping clarification now explicitly requires the registered timer to check
captured context before sync/overlay maintenance and stop rescheduling when stale,
even before a click rejects that form. Healthy scheduling remains; internal timer
dispatch creates no user-control activity. This inherits the approved context and
observation rules without a new permission or catalog version. Repair is confined
to `frmShipmentsTally`; removing its timing-start pass-through retains the existing
typed timer and eight packaged timing checks while preserving the module limit.
`Slice4beShippingInterruptions.ps1` protects the real callback and UI-yield paths.
This RED precedes the focused GREEN and candidate gates recorded below.

The isolated `deploy/validation-shipping-timer-context` repair now records
**322 PASS /46 missing-activity FAIL** across the same368 checks, preserving every
prior GREEN. The three timer failures and all30 interruption/timer checks pass.
Build/all-five compile/cold-start and static contracts pass; the form shrinks by
three lines and all28 module limits hold with unchanged dynamic/duplicate counts.
The first attempt stopped before Shipping with RPC failure and Excel
c0000028/ntdll.dll; the unchanged pinned retry supplies the focused GREEN, not a
native-stability clearance. Visible comparison with the preceding guard candidate
retains the legible rejection notice and the same key-editor/header crowding.
Candidate release-gate results follow. See the linked context evidence for exact
reports and scope; this does not close broader4be.1 or4be.2-4be.6.

Smoke passes86/86, but live-role stops39 PASS /1 harness failure at Production
Complete Run with c0000028/ntdll.dll. The preceding pinned guard package set
reproduces39/1 at that same step and fault; this is also observable without the
timer repair, not a cleared native defect. Independent gates pass full-chain30/30,
Viewer, launchers3/3 and Shipping layout1/1. A combase.dll/c0000005 fault within
the passing chain still prevents native clearance. The fixed-phase diagnostic
reproduces39/1 and reaches P08 before quiet-UI entry/typed Production completion,
but never P09 before output restoration; narrow that interval before a runtime fix.

Reusable Production first fails fixture sign-in (0/2, not behavioral RED). Its
full-settings fresh retry signs in, then fails0/1 with RPC/native c0000028 at the
batch-scale contract. No candidate reusable Production/restart GREEN is claimed.
All test handles are terminal, Excel is closed and40 package/eight current source
pins match. The timer is a focused GREEN checkpoint; live-role/native/Production
acceptance and the full4be activity/Settings/Viewer/Action Path scope remain open.
Code checkpoint **0aa0084** is committed and pushed. Preserve the pinned candidate
and all existing GREEN evidence while narrowing the native failure; do not label
this checkpoint a completed Slice4be or substitute it for user acceptance.

**Unchanged-source native/gate comparison, 2026-09-12:** Fresh isolated
`validation-shipping-timer-rebuild` builds/compiles all five packages and passes
cold start; all170 exported component hashes match the earlier timer candidate.
Ordinary live-role48/48 and full reusable Production/restart2/2 pass. Shipping
retains322 PASS /46 missing-activity FAIL, all368 check identities and every prior
GREEN. This same set passes smoke86/86, ordered full-chain30/30, Viewer and Shipping
layout1/1. Combined launchers fail1 PASS /1 harness exception at Production batch
scale after successful sign-in, with ntdll.dll/c0000028. A separate combase.dll/
c0000005 occurs during the passing full chain. Native stability and combined
launcher acceptance remain open; do not combine package sets or traced passes to
claim otherwise. Expanded unsaved Complete Run tracing passes48/48; an identical
service-source-only reload still fails39/1. Neither establishes a runtime cause.

Test maintenance rejects a failed fixture sign-in before workflow callbacks and
updates the stale source assertion to follow Shipping's existing typed context/
bound-form factory. The ordinary full Production2/2 retains its complete scope;
the source checker passes24/24. Runtime remains0aa0084 with no architectural change.
The fresh capture retains the legible context notice and prior key/header crowding.
See [native validation evidence](../../invSys_fork/tests/integration/plan022_slice4be_native_validation_results.md)
for exact variant/gate reports and limitations. All4be.1-4be.6 work and Release1
acceptance requirements remain intact.

**Shipping workbook-close test entry, 2026-09-12:** Continue4be.1 under D18's
captured-workbook lifetime and non-click rules. The supplemental
`Slice4beShippingWorkbookClose.ps1` closes the real captured disposable workbook
with normal Excel events enabled while an unrelated workbook is active, then
invokes the registered late timer. Assert released form/callback bindings, no
reopened workbook or owner/submission entry, no fabricated activity, and preserved
saved staging/exact keys/unknown values, authority bytes and unrelated workbook.
Unsaved facades inspect binding state only; they do not call the close handler,
suppress events or change runtime logic. A surviving binding or redirected action
would be behavioral RED; already conforming lifetime behavior needs no runtime
repair. This test does not claim a forcibly retained stale form or capability-loss
case. Existing368 checks and all322 GREENs remain required.

The calibrated normal-event run passes **333 checks /46 missing-activity failures**,
379 total, including11/11 workbook-close checks and every prior check/GREEN with
no duplicates or other failures. The initial close fixture stopped319/48 because
the shared harness starts with events disabled; normal Shipping startup/event
initialization corrects that setup. A subsequent native67/1 failure before Shipping
and an Excel-open refusal are retained separately. No runtime close repair was
needed. All test processes are terminal, Excel is closed and45 package/eight
current source pins match. Static runtime metrics and28 module limits are unchanged.
The linked native validation evidence records exact runs and scope. Continue
capability-loss and pending/uncertain owner outcomes before Shipping's D18 catalog/
activity implementation; full4be.1-4be.6 and native/human acceptance remain open.
Code checkpoint **97c74eb** is committed and pushed. This checkpoint changes tests,
maintenance metadata and evidence only; runtime remains0aa0084 and no Slice4be
completion or native-stability repair is claimed.

**Shipping capability-loss D13 entry, 2026-09-12:** Continue4be.1 under D18's
existing role-capability and owner-fact rules. `Slice4beShippingCapability.ps1`
starts from genuinely staged active/held rows and a live authorized Shipping form.
Calibrate all seven mutation handlers at their owner entry, revoke only SHIP_POST
in the disposable Auth fixture, and verify Core denies that capability while the
same session remains signed in. Repeating the handlers must reject visibly before
mutation. Stopped-owner probes distinguish handler entry from owner authorization;
a separate real Hold action proves whether local staging is actually protected.
Expected behavioral RED is unauthorized local movement or missing pre-mutation
denial, not an Auth/fixture failure. Preserve all379 prior checks/333 GREENs,
captured workbook, exact keys, unknown values, canonical bytes and unrelated work.
No runtime edit or new capability is introduced in this test entry.

D18's Shipping mutation permission constraint now explicitly applies the existing
SHIP_POST requirement before local mutation owners as well as submission. Recheck
after yields/between separate row mutations and retain captured context around
Core authorization. The fixed notice is **Shipping permission could not be
verified. Review Shipping access before continuing.** A failed check cancels the
form's pending sync without implying a specific Auth/Config cause or rollback.
Core permission/security ownership, service checks, timer's existing context-only
entry contract and launcher/Close behavior remain. This semantic-inheritance
constraint adds no permission or activity catalog and is recorded before repair.
The bounded Shipping context helper can own the combined check; the form's
existing guard delegates to it without exceeding its current3019-line limit.
Focused RED must precede that implementation.

Unchanged-package RED is374 PASS /61 FAIL across435 checks, with all379 prior
checks/333 GREENs retained. Core denies SHIP_POST while the session stays signed
in; all seven handlers still enter mutation owners without a permission notice,
and real Send Hold moves staging. These15 failures are distinct from the46 missing
activity failures. After normative clarification0cc61c9 is committed/pushed, the
bounded candidate passes389/46: all56 added capability checks and every prior
GREEN. It changes only the Shipping form/helper, keeping the form3019 lines and
adding17 helper lines. All five build/compile/cold-start checks pass, with exactly
two changed exported component hashes. Dynamic/duplicate counts and all28 prior
module limits hold. The permission notice capture is legible; existing key/header
crowding remains. Smoke86/86 passes; candidate role/chain/Viewer/launcher/Production/
layout gates remain in progress. See [capability evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_capability_results.md).
This is focused permission GREEN, not Shipping activity or Slice4be completion.

The same candidate passes full-chain30/30, Viewer, combined launchers3/3, full
reusable Production/restart2/2 and Shipping layout1/1. Initial live-role stops14/1
at Receiving ConfirmWrites with DISP_E_EXCEPTION and no observed Windows native
fault for that run. Its remaining empty Excel instance exits after Quit. A fresh
ordinary live-role run then passes48/48, retaining the initial failure separately.
An Excel combase.dll/c0000005 fault during the passing full-chain window remains
unresolved. All50 package/eight current source pins match; Excel is closed and
all handles are terminal. Static schemas,24 launcher source checks and28 preceding
module limits pass. The linked evidence gives exact run/capture provenance.
Continue permission loss during yields, unavailable access evidence and pending/
uncertain source outcomes before completing the comprehensive Shipping activity
contract; native stability, human comparison and full4be.1-4be.6 remain open.
Code checkpoint **22e14b6** is committed and pushed after normative0cc61c9.
It is the focused permission repair with the candidate evidence above, not a
native-crash fix or full Shipping activity acceptance.

**Shipping access-interruption test entry, 2026-09-12:** Continue4be.1 against
unchanged22e14b6 packages under D18's permission/yield and D5 fail-closed rules.
`Slice4beShippingAccessInterruptions.ps1` revokes the generated fixture's SHIP_POST
at the actual Add/To Shipments/Shipments Sent pending-status yield, without replacing
Core authorization. Separate cases temporarily move the fixture Auth or Config
file aside before each of the seven mutation handlers. Require denied/unavailable
Core permission with the same signed-in session, no mutation-owner entry, exact
fixed access notice, no recreated authority file, and preserved staging/keys/
unknown values/canonical bytes/unrelated workbook. Restore fixture files and verify
their bytes. Calibrated owner probes protect dispatch; they do not establish
completed business outcomes. All435 prior checks/389 GREENs remain required.
This adds protecting tests under the existing approved contract; conforming cases
may pass without a runtime edit. Any behavioral failure must be distinguished from
instrumentation or fixture failure before proposing a repair.

**Auth creation discovery / pending D8-A decision:** The first interruption run
passes the three actual-yield revocations and rejects Shipping's first missing-Auth
mutation, but recreates Auth during its ordinary permission read. Cleanup then
stops rather than overwriting that unexpected file; subsequent sign-in masks the
cleanup exception. Preserve that incomplete run. The helper now preserves each
unexpected fixture file under a unique sibling name, restores the original, and
does not retry sign-in after an incomplete helper. This permits the remaining
unchanged-package cases to run without suppressing the creation observation.

The test entry's blanket no-recreation expectation was too broad to cite D5 for
Auth: D5 explicitly governs Config, while the normative Phase 6 auto-bootstrap
entry includes Auth. Auth creation is an observed behavior and proposed D8-A RED,
not an already-approved Auth read-contract breach. Architecture v4.11 now contains
the concrete **pending D8-A Auth read/provisioning separation** proposal. It would
restrict ordinary reads to existing exact-target Auth, with no creation/repair,
and retain explicit authorized setup. Approval must precede any conflicting Core
implementation. This discovered decision does not narrow remaining 4be coverage.

The corrected unchanged-package run completes **505 PASS/53 FAIL across558
checks**: all435 previous identities/389 GREENs remain, with no duplicates or
unexpected failures. All123 additions execute:116 PASS and seven Auth no-creation
findings awaiting the D8-A decision. Each missing-Auth handler stops mutation;
all actual-yield revocations and missing-Config cases pass. All files are restored
with matching protected bytes, and the existing actual workbook-close checks pass.
The46 missing-activity failures remain. Fifty package/eight source pins match;
Excel is closed and focused/static sessions are terminal. Regenerated static
metrics and28 preceding module limits hold. See [access-interruption evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_access_results.md)
for exact runs, the retained cleanup failure and the distinction between current
permission GREEN and proposed Auth-contract RED. No runtime or broad-gate change
is claimed; approval of D8-A remains pending while submission/activity work remains.
Code checkpoint **5322f9f** is committed and pushed. It contains the test-only
extension and sanitized evidence; runtime remains22e14b6. D8-A is still a proposal,
not an effective architecture change or an authorized Core implementation.

**Shipping submission discovery test entry, 2026-09-12:** Continue 4be.1 under
D18's owner-fact and exact-reference rules while D8-A remains pending. The
unchanged packaged Add handler now gets a separate generated fixture after all
558 preceding checks. Unsaved observers retain the actual commit result and
reserve ID. Four controlled public submission faults cover unavailable server,
lost acknowledgment with successful local fallback, exception after actual server
acceptance, and lost acknowledgment with unavailable fallback. Successful writes
and all authorization owners remain real; fault hooks never fabricate an event ID
or a successful write. Observe both public calls, exact ID propagation and whether
the Domain log has applied that specific ID. Normal Remove clears successful
local staging between cases; the uncertain case runs last in its isolated fixture.

The new helper must calibrate before source-outcome assertions become acceptance
evidence. It does not register Shipping controls or implement activity. Existing
activity tests currently derive reference expectations from applied inventory logs;
discovery must determine whether this omits pending submissions before those tests
can protect comprehensive observation. A compiler/instrumentation failure remains
harness failure, not meaningful D13 RED. Preserve all558 prior checks/505 GREENs,
the46 missing-activity assertions and the seven separately pending D8-A findings.

Initial submission discovery completes605 checks,552 PASS/53 FAIL, preserving all
558 prior identities/505 GREENs and passing47 new source/owner checks. Exact IDs
survive negative and exceptional acknowledgment; even owner failure may retain a
server-accepted ID absent from the Domain log. This proves that applied-log-only
activity expectations miss pending submissions. The first helper also omits
reestablishing fixture paths after unsaved Core edits, creating a default template
and fixture operator outside its intended temporary tree. The template is isolated
under ignored evidence; the external fixture is left untouched after automatic
approval review rejects combined cleanup. Corrected explicit template/operator
roots and an unsaved temporary local-staging root now protect isolation, with
persisted-ID read-back and source-byte checks added.

Two full runs of that correction stop67/1 at the existing Shipping setup before
the new helper, both with native ntdll.dll/c0000028. Retain both failures and the
empty recovery-instance cleanup evidence. A separate `-ShippingSubmissionOnly`
diagnostic route now permits focused calibration without claiming full-matrix
GREEN or a native repair. See [submission evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_submission_results.md).
The default full route and all prior assertions remain; D8-A is still pending.

The focused diagnostic completes **136/136**, including66 submission checks and
all47 initial submission GREENs. Persisted server/local identities, exact owner
references, absent Domain application, bounded fixture paths and source/template
byte preservation pass. All50 package/eight source pins match; runtime stays
22e14b6. Static schemas and28 preceding module limits pass with unchanged runtime
metrics. No native fault is observed in that focused window, but the revised624-
check full route and earlier native failures remain unproven. Do not combine the
605-check initial run and136-check focused run into a full GREEN. Separately
verified single-file moves safely isolate the initial generated operator and its
fixture-only staging file after the rejected combined cleanup; unrelated files
and surrounding folders remain untouched. Correct the normal activity test's
reference expectations from actual owner submissions next, then refine the
Shipping catalog/outcome contract under D18 before implementing observation.
Code checkpoint **01cb839** is committed/pushed with the submission tests,
sanitized evidence and regenerated maintenance reports. Runtime remains22e14b6;
the current full route's native failures, missing activity and D8-A decision stay
open. The focused pass is evidence for source semantics and fixture isolation only.

**Shipping owner-reference test correction, 2026-09-12:** Continue 4be.1 by
replacing the normal activity test's applied-log expectation with exact identities
acknowledged by `QueueShippingPayloadEventServerFirst` during that actual handler.
An unsaved observer records only real positive returns; reset it per user action,
check unique valid IDs and submission-count agreement, and keep Domain-log reads
as separate application evidence. The historical ExactAppliedSourceReferences
check name remains for identity continuity, but its expectation now includes all
owner-submitted IDs, including pending work, and excludes earlier actions merely
processed during catch-up. The ordinary fixture expects one source for Add,
Remove, AddAgain and Send, and none for its delta-only Update, Hold, Return and
already-reserved Stage. This is a test correction to D18, not a new runtime contract.

The first corrected full attempt again stops67/1 before these checks. Run's
reported seven arguments include the macro name: source call order identifies
the six-value `BootstrapWarehouseLocalAdmin` fixture call, not the zero-value
Shipping launcher. The normal helper sets its bootstrap paths after that call;
the passing focused helper sets them before generation. Read-only unsaved getter
instrumentation now observes bootstrap-root matches before/after the Operations
test edits, and explicit roots are verified before generation. Failure diagnostics
log only the macro name and argument count, never values. This calibrates fixture
setup and tests a bounded explanation; it does not claim a native runtime repair.
All prior assertions, pending D8-A findings and unresolved native evidence remain.

**Owner-reference checkpoint, 2026-09-12:** The corrected complete route finishes
650 checks:597 PASS/53 FAIL, preserving all605 preceding check identities and all552
preceding GREENs, with45 additional passing checks and no duplicates. All66 source
submission checks now pass within the full route. The remaining46 missing-activity
checks and seven pending D8-A findings retain their separate status. Send submits
one source while applying four, including three earlier actions; Add and Remove
retain acknowledged pending IDs. The new expectation therefore prevents both
missing pending references and misattributing catch-up work to a later click.

Bootstrap-root observations are True before and after Operations instrumentation
and before generation. They do not support the root-reset hypothesis. No Excel
Application1000 fault is observed in the completed-run window, but neither the
passing setup nor explicit root setters establish a native repair. The preceding
67/1 run records ntdll.dll/c0000028; it remains a harness/native failure, not product
RED. Runtime22e14b6 and all50 package/eight source pins are unchanged; Excel is
closed. The [sanitized owner-reference record](../../invSys_fork/tests/integration/plan022_slice4be_shipping_owner_reference_results.md)
contains exact scope, counts and evidence. Next register/refine Shipping activity
outcomes and references under D18, preserving these650 checks before implementation.
This checkpoint does not close4be.1, approve D8-A or establish full-release acceptance.
Static evidence is regenerated with all three schemas and 28 previous module limits
passing; procedure/dynamic-call/duplicate counts are unchanged. Both edited
PowerShell parsers, 52 local Markdown targets and diff/status review pass.
Code checkpoint `3a6f345` contains the test correction and sanitized evidence;
runtime remains `22e14b6`.

**Shipping catalog 8 D13 entry, 2026-09-12:** Continue 4be.1 under the normative
D18 Shipping activity/source-reference clarification. Register SHIPPING_ADD,
SHIPPING_UPDATE, SHIPPING_REMOVE, SHIPPING_HOLD, SHIPPING_RETURN, SHIPPING_STAGE
and SHIPPING_SEND under SHIPPING_WORKFLOW/SHIP_POST. Preserve catalog versions 1-7.
The exact outcomes distinguish local STAGED, accepted PENDING, Send processing/
refresh CONFIRMED with unknown Domain effect, and FAILED with mixed per-source
Submitted/Unknown facts. This inherits the approved D18 contract; it does not
change business behavior, approve D8-A, or claim comprehensive handler coverage.

Extend the existing 650-check packaged Shipping route with exact catalog/outcome
and reference-validation checks, retaining actual handler RED. Expect missing
catalog 8 definitions and rejection of legitimate Shipping source references.
Implement Core catalog/reference support only after that RED; record its GREEN
separately from remaining handler failures. Expected source changes are
modActivityCatalog, modActivityReferences and a focused Shipping outcome module.
Follow with owner-fact plumbing and actual handler observation, covering partial
Update/Remove/Stage, Send processing, tracking/policy failures and role regressions.

Source review also finds a ROW compatibility slot in the reachable local Shipping
persistence writer (PersistShipmentRowsLocal -> HoldRowField). D14 forbids that
compatibility field. Preserve the finding for a focused save/reopen identity and
unknown-column test; do not bless it through activity registration or assume the
existing event-payload checks cover local persistence. No persistence change is
made by this catalog work.

**Shipping catalog 8 focused result, 2026-09-12:** Normative commit `37825b5`
precedes the Core implementation. Packaged RED executes 204 supplemental checks,
115 PASS/89 FAIL for absent definitions/source eligibility. The candidate at
deploy/validation-shipping-catalog-eight passes all 213 catalog/reference/policy
checks, retaining all 204 RED identities and adding nine passing catalog-7 policy
regressions. Only modActivityCatalog/modActivityReferences and the new
modShippingActivityCodes differ in exported compiled source. All five builds,
explicit compiles and Operations cold start pass; package smoke passes 86/86.

Both full invocations then stop in the six-value Admin fixture bootstrap before
Shipping handlers execute (RED overall183/90; candidate overall281/1). Both record
ntdll.dll/c0000028. This establishes catalog GREEN only, not preservation of the
full650 handler route on the candidate. Independent live-role validation stops
39/1 during the Production completion portion with the same native fault; its
failure remains recorded while the independent full-chain and remaining gates run.
The full Slice4be and release goal remain active, with native stability unresolved.
See the [Shipping catalog evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_catalog_results.md)
for exact runs and scope. No Shipping handler activity is implemented by registration.

**Catalog candidate gate closure, 2026-09-12:** The ordered full chain stops4/1
with ntdll.dll/c0000028. Viewer returns FAIL after fixture sign-in fails; its
verified information dialogs are captured/acknowledged without changing sign-in,
and the report retains SignedIn=False. Combined launchers reach Receiving then
stop1/1 with another native fault. Separate full reusable Production/restart passes
2/2 and Shipping layout1/1, with no Application1000 Excel fault observed in those
two windows. The queue ends failed overall; no full-release or native acceptance
is inferred from the independent passes. The candidate's full Shipping and full
Receiving activity regressions remain required before broad acceptance.

Static regeneration passes three schemas and all28 previous module limits;
178 components/5514 procedures/123883 lines retain8 literal/45 unresolved dynamic
calls and195 duplicate bodies. Both edited PowerShell parsers,46 local Markdown
targets and diff/status checks pass. All55 package pins, six unchanged source pins
and three current Core pins match; Excel is closed and all handles are terminal.
The new verify-catalog-eight-pins.ps1 supersedes only the two intentionally changed
Core expectations in the older verifier. Preserve unrelated user handoff067 and
critique023. Next establish a valid candidate handler-test route and add owner-fact
plumbing/observations under the registered contract; do not repeat failed gates
unchanged or claim source-reference registration creates user activity.
Code checkpoint `469703f` contains the catalog/reference implementation, focused
tests and sanitized evidence. Its isolated candidate is retained for diagnosis;
accepted deployment is unchanged and Slice4be remains incomplete.

**Validation recovery entry, 2026-09-12:** Continue 4be.1 without a runtime contract
change. The Viewer harness previously continued into public actions after failed
fixture sign-in, causing repeated information dialogs. Add fail-fast setup with
only an allowlisted status and credential-match/entry booleans; a deliberately
rejected fixture credential must stop before Viewer entry. The first diagnostic
normal run passes the full Viewer gate, but does not explain the earlier sign-in
failure. D8-A remains pending; no Auth implementation changes are authorized here.

The early Operations-probe installation experiment stopped281/1 at bootstrap
with ntdll.dll/c0000028 and was reverted. Explicit unsaved fixed-phase bootstrap
tracing then completed the full863-check route:810 PASS/53 FAIL, preserving all650
preceding check identities and GREENs plus213 passing catalog/policy checks. The
trace does not identify a native cause or prove an uninstrumented repair.

Both normal Viewer runs pass; deliberately rejected sign-in stops before public
entry with AUTH_STATUS_4. A subsequent Excel Application1000 fault (c0000409,
module unknown) remains recorded for the rejection window; clean native shutdown
is not proven by the setup guard. No Excel fault was observed in the two normal
Viewer windows or the full traced Shipping window. All55 package pins and nine
source pins match; Excel is closed. Runtime/architecture remain unchanged, and
D8-A is still pending. See [validation recovery evidence](../../invSys_fork/tests/integration/plan022_slice4be_validation_recovery_results.md).
Next add typed Shipping owner facts and real handler observation, preserving all
810 current GREENs and the full release/visible acceptance obligations.

**Shipping exact-owner-evidence D13 entry, 2026-09-12:** Extend the actual-handler
test before VBA changes with27 assertions across eight established normal actions
and invalid quantity: exact registered ControlId/owner/surface, REQUESTED facts
with empty references, and exact outcome/EventCode/severity/effect. Independent
fixture expectations are Add/AddAgain/Remove PENDING, delta-only Update/Hold/Return/
already-reserved Stage STAGED, completed Send processing/refresh CONFIRMED, and
invalid quantity REJECTED. Source acknowledgment and owner-state checks remain
independent. Expect absent activity RED on the unchanged catalog candidate; retain
all863 preceding checks. This tightens existing D18 evidence, without changing
architecture, granting D8-A approval or inferring per-source Domain application.

**Exact-owner-evidence result, 2026-09-12:** The first strengthened invocation
with bootstrap tracing stops281/1 with another ntdll.dll/c0000028 before the new
assertions. A bounded `-ShippingBeforeSharedFormsForTest` diagnostic instead runs
the full Shipping group before the shared Settings/Production form exercises,
then retains all remaining checks, without bootstrap tracing or package changes.
It completes890 checks:810 PASS/80 FAIL. All863 preceding identities and GREENs
remain; the27 added exact-outcome assertions are meaningful missing-activity RED.
The other failures remain46 missing activity and seven pending D8-A findings.
No Application1000 Excel fault is observed in this run's window; no general native
repair is claimed. All55 package/nine source pins, three static schemas,28 prior
module limits and unchanged maintenance metrics pass; Excel is closed. See
[exact Shipping outcome evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_exact_outcomes_results.md).
Next implement the typed owner facts and actual observations already required by
D18, including partial/mixed failures, without another unchanged diagnostic retry.
Code checkpoint `f2ba51f` contains the test changes and sanitized evidence; runtime
remains `469703f`. Four PowerShell parsers,53 local Markdown targets and reviewed
diff/status checks pass. Unrelated user handoff067 and critique023 remain untouched.

**Shipping handler implementation entry, 2026-09-12:** Under unchanged D18,
add per-action typed Shipping owner facts and handler-only observation. Owners
retain every exact accepted/uncertain source identity and report local completion,
required-step failure and completed Send processing/refresh separately. No report
text determines activity success, and a later row cannot erase an earlier failure.
The public handlers observe before authorization/validation after captured-context
validation; optional tracking failure preserves business behavior. Changes are
limited to Shipping form/posting modules and focused fact/observation helpers in
Operations. Cohesive text/selection helpers are extracted to retain existing form
and main-module size limits. Protecting RED is the full890-check810/80 result;
partial/fallback, policy/tracking, packaged and full release gates remain required.
No D8-A or business/persistence contract change is implemented by this entry.

**D18 assertion reconciliation, 2026-09-12:** The first owner-activity candidate
builds/compiles and completes890 checks at873/17. All73 missing normal activity
assertions turn GREEN. Ten old assertions conflict with already-approved D18:
three NoCrossContextActivity checks count the legitimate pre-interruption REQUESTED
record as forbidden; seven FixedAccessNotice checks prohibit the independently
required tracking-unavailable notice when Config cannot be read. D18 now explicitly
clarifies these existing rules, without allowing post-loss attribution or changing
the fixed primary access notice. This is semantic inheritance, not approval of a
new authorization, audit or persistence contract. D8-A remains pending.

Preserve the ten assertion identities and their intended invariants. Capture actual
record contents immediately before the test's real sign-out to prove later reads
contain only the same pre-loss attempt. Require the exact primary access line and
only the fixed Config tracking notice as an optional second line. Add separate RED
assertions for the required pre-loss attempt and tracking notice, plus actual
fallback/uncertain-submission and unavailable-store handler coverage. Retain the
first873/17 result; do not retrospectively label it full GREEN or erase its evidence.

**Shipping owner checkpoint, 2026-09-13 (code `ec8110a`, pushed):** The implemented Operations candidate
retains the original873/17 evidence. Expanded prior-package testing completes953
checks at828/125:117 missing activity assertions, seven pending D8-A findings and
one test-only SHA-256 formatting mismatch, independently calibrated and corrected.
All890 preceding check identities remain. The diagnostic prepares the same three
Admin-generated fixtures before Shipping probes and consumes each once; it does
not skip bootstrap or explicit Seed. Candidate attempts still fail in Admin Seed
(5/1, then1/1 with a removed Seed-order experiment), so expanded GREEN is pending.

All five candidate builds/compiles, smoke86/86, Viewer and combined launchers3/3
pass. Live-role remains39/1 with a native failure. Full chain reports30/30 and its
ordered child48/48, but the window includes an Excel native fault and source review
finds that the runner ignores the ordered child's exit code. Retain the assertions;
full-chain acceptance remains unproven. The runner's actual child-decision test
records2 PASS/2 FAIL before adding OrderedLiveProcessCompleted and4 PASS/0 FAIL
afterward. The corrected packaged chain completes31/31, preserving all30 prior
identities, but again records Excel combase.dll/c0000005 in its window. The child
exit requirement is proven; native-clean full-chain acceptance remains open.
Independent reusable Production/restart2/2 and Shipping layout1/1 also pass.
Maintenance preserves all28 prior module limits and8/45/195 dynamic/duplicate
counts; the16 extracted Shipping helpers preserve their bodies exactly. All60
package pins and14 current source pins match after the runs; Excel is closed.

Next required source-reference case: actual Add with both submission routes
refused before writing. Core allocates the ID before inbox availability/schema
checks; a nonempty ID alone cannot establish an attempted write. D18 prohibits
allocated-only references. Prove that failure before refining the primitive owner
facts, and keep partial/mixed outcomes, policy variations, remaining coverage,
release/visible gates and the separate D14 TSV correction open. D8-A is still
pending explicit approval. See [owner activity evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_owner_activity_results.md).

**Pre-write identity D13 entry, 2026-09-13:** The existing candidate's real Add
now reproduces the D18 breach:402 checks complete401 PASS/1 FAIL. Both Core
routes allocate/preserve the same ID and refuse before either actual writer;
independently calibrated write counters remain zero and authority/submission
files remain unchanged, but the activity result includes that unsubmitted ID.
The only RED is Shipping.Submission.PrewriteRefusal.Activity.ExactAppliedSourceReferences.
An earlier280/1 probe-anchor failure is retained as harness evidence, not RED.

Follow the normative submission-entry refinement: add optional primitive
writeAttemptedOut facts to the two Core queue routes and preserve both facts in
Shipping, including exceptions. Keep the allocated ID and Boolean business
contracts unchanged. Use coherent helper extraction to retain the Core/Shipping
module-size limits. Protect accepted writes, lost/exceptional acknowledgments,
known pre-write refusal, exact identity, bytes and unknown columns through the
same packaged handler route, then rerun required package/release gates. This
corrects existing D18 behavior and does not approve D8-A.

**Pre-write identity candidate evidence, 2026-09-13:** Code807c7e1 is committed
and pushed under normative590d41a. This remains a partial Slice4be checkpoint;
the focused handler route passes402/402 with all401 previous GREENs and check
identities preserved. The complete prepared-fixture Shipping route completes974
checks at967/7; only pending D8-A recreation findings fail. Both890-check baselines
and the953-check expanded RED retain every identity/GREEN, with no duplicate
checks. No Excel fault is observed in these focused/full Shipping windows. Five
builds/explicit compiles and smoke86/86 pass. All28 previous module limits and
8/45/195 maintenance counts remain; five extracted helper bodies are unchanged.
An independent read-only comparison explains the form compiled-hash difference
as exactly four trailing blank lines, preserving both package files. Viewer,
independent Production/restart2/2 and Shipping layout1/1 pass. Separate live-role
39/1 and combined-launcher1/1 failures coincide with Excel native faults and remain
open. Sixteen source-harness import lists now include the extracted JSON helper;
all parse. The first full-chain attempt was a missing-dependency compile setup
failure. After repair, Create Warehouse passes15/15, but full-chain stops at4 PASS
and1 Harness.Exception with RPC0x800706BE before ordered live-role completion.
That RPC window also contains an Excel fault. Receiving stops at591 PASS/30 FAIL
across621 checks:29 outdated catalog3-7 assertions and one keyboard-input harness
failure. Its preserved845/845 comparison leaves225 identities unexecuted. Both
Receiving test allowlists now include approved8; focused navigation stops106/1
at the same first keyboard action. Neither Receiving window shows an Excel fault.
Full Receiving preservation remains open. Next calibration must prove the native
keyboard target is within the intended form after another workbook activates;
same-process ownership alone is not that proof. Final65 package/16 source pins
match with Excel closed. See
[pre-write identity evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_prewrite_refusal_results.md).

**Receiving native-target follow-up, 2026-09-13:** Codea300cc3 is committed and
pushed. Focused266/266 passes after
calibrating the actual native focus transition and repairing test input routing.
The protecting RED is111 PASS/2 FAIL across113 checks; its111 GREENs remain.
Nine new checks protect wrong-window refusal and the exact captured object/native
form across workbook closure. Native actions still enter the real handlers and
Core readers; no runtime form, authority rule or package changes. The fixture
explicitly shows Excel for mouse input and restores visibility. Native focus may
activate only the unchanged workbook or exact captured object. The final stale
form is hosted independently of its captured business workbook so actual input
can reach its missing-workbook guard. Full Receiving854/854 now retains every
prior845 check and GREEN, with nine added checks and no duplicates. A fresh form
capture was inspected;65 package/16 source pins match with Excel closed. No Excel
fault is observed in these calibration/focused/full windows. Runtime maintenance
counts and all28 module limits remain; all three report schemas pass. Separate
native failures, D8-A approval and comprehensive Slice4be acceptance remain open.
See [native-target evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_native_target_results.md).

**4be.2 Settings surface RED, 2026-09-13:** Code `893f5e0` is committed and pushed.
Product-facing Settings work is now
active alongside the remaining 4be.1 coverage. The separate packaged D5 route
with `-CheckTrackingSettings` completes 26 checks: 21 PASS / 5 FAIL, retaining every
18/18 prior D5 identity and GREEN. The actual initialized/shown form lacks the
approved tabs, three tracking sections, separate actions, capture-off control
and personal view selector. Config bytes are unchanged by opening. This is
surface RED only; tracking save/profile/preference behavior and Ribbon-open
authorization are not yet proven. An initial observer-installation error and
native fault are retained separately, not counted as product RED. The corrected
run has no observed Excel fault and its form capture was inspected. No runtime
or architecture changes; 65 package / 16 source pins remain. See
[Settings evidence and required next gates](../../invSys_fork/tests/integration/plan022_slice4be_tracking_settings_results.md).

**4be.2 tabbed UI checkpoint, 2026-09-13:** Code `3ea260f` is committed and pushed.
Admin Settings now has General and
Event Tracking pages in `deploy/validation-tracking-settings-tabs-final`.
Existing editors/handlers stay on General; Close/status remain shared. The three
tracking sections explicitly show that their editors are unavailable. This is
an intermediate implementation, not completion of the approved Settings scope.
Expanded old-package RED is 23 PASS / 8 FAIL across 31 checks; the final candidate
is 28 PASS / 3 FAIL with all prior identities/GREENs retained. Focused tab/layout
checks pass 10/10; D5 remains 18/18. Both displayed pages fit after a measured
footer overflow was corrected by increasing form height to 696 points. All five
packages build/compile, Operations cold start passes, and packaged smoke is
86/86. Final General save/status and Event Tracking captures were inspected.
Save actions, capture control and personal selector remain RED. Policy/profile
persistence, preference isolation, new page activity coverage and all broader
acceptance gates remain open. See the updated Settings evidence above.

**4be.2 tracking policy checkpoint, 2026-09-13:** Code `be0e8a0` is committed
and pushed. The isolated
`deploy/validation-tracking-policy-save` candidate implements staged Admin
collection/visibility/sequence flags, capture and Admin visibility flags, and
warehouse default view. Its Save/Reset/Reload actions use the captured context;
headless Core validates the whole policy and expected version, preserves unknown
columns, appends metadata/control rows and saves once. This implements existing
D18/D5 without changing the normative contract. The no-write scaffold's real
Save action gives meaningful RED: 47 checks, 44 PASS / 3 FAIL, including failure
to publish a version. The expanded candidate gives 56 checks, 54 PASS / 2 FAIL,
with all 25 focused policy checks and 18 D5 checks passing. The two remaining
broad failures are profile/preference actions and personal choices. Shared
activity remains 70/70; both candidates build/explicitly compile all five
packages and pass Operations cold start. The saved version-2 capture is inspected.
Full Receiving passes 854/854 with every prior check identity/GREEN retained;
packaged smoke passes 86/86, and source-harness import/test 1 passes 1/1.
Static dynamic-call/duplicate metrics are unchanged and all 28 prior module
limits hold. All 75 prior package pins and 16 preserved runtime source pins match;
10 policy candidate package hashes are pinned. Excel is closed. Eleven bounded
windows show no observed Excel Application 1000 fault; historical faults remain.
New Settings action coverage and policy-save observations, further policy edge
cases, detail profiles, personal preferences, Operations Settings without Admin,
recording, Viewer/comparison and broader release/user gates remain open. See
`tests/integration/plan022_slice4be_tracking_settings_results.md` in the code repo
for exact test runs, limitations and continuation. Do not treat this checkpoint
as completed 4be.2 or use it to weaken the acceptance table below.

**4be.2 cancelled-policy-save correction, 2026-09-13:** Code `f629604` is
committed and pushed. A real Excel
WorkbookBeforeSave cancellation leaves Config unchanged but the prior candidate
incorrectly reports success and reloads away staged edits. Expanded packaged
RED is 65 checks: 61 PASS / 4 FAIL. Core now verifies Workbook.Saved after its
single save call and uses the existing unverified-save cleanup on cancellation;
the form retains staging. The corrected candidate is 63 PASS / 2 FAIL, with all
34 policy and 18 D5 checks passing. Added real per-control, Reload and Close
action tests pass. Only modTrackingPolicyCommand changes across 180 compiled
components; all five compile, cold start passes and smoke is 86/86. Earlier
activity 70/70 and Receiving 854/854 retain their previous-candidate scope and
were not rerun for this command-only correction. D18/D5 is unchanged. Proceed
with Event Detail profiles/personal preferences; retain policy edge cases,
new-control coverage, policy-save observations and broader acceptance as open.
The Settings evidence record names exact RED/GREEN reports and limitations.

**4be.2 Operations Settings continuation, 2026-09-13:** Protect the Viewer Settings
entry and real form handlers with a four-package directory physically omitting
Admin. D18 specifies an Operations-owned captured modeless instance, read-only
policy projection and personal controls using the existing Core boundary. Reuse
must retain staging; stale context must not retarget; closing Viewer closes its
Settings. Preserve the 145-check Admin/preference baseline, Viewer state and
authority bytes. Surface and action RED precede implementation; activity coverage
for the discovered controls and broader release gates remain required.
The current candidate passes **187/187** combined Settings checks, including
36 Operations and six actual Admin Close checks with every preceding identity
and GREEN retained. Operations runs with four isolated packages and named Admin
absence. Populated Viewer preservation, default/native maximize/restore and
readable aligned-header captures pass. Five packages compile, cold start passes
and packaged smoke is 86/86. Full-chain retry passes 31/31 with ordered live-role
child 48/48; the initial projection-rebuild native fault remains recorded. All
115 prior / 25 new package pins and 16 protected source pins match. Exact evidence,
the narrower populated Viewer run's scope and prior failed harness runs are in
the maintained Settings results. Proceed with comprehensive Events, Viewer
profile/preference consumption, both Action Path presentations and new-control
activity coverage; human acceptance and the full Slice 4be outcome remain open.

**4be.2 discovered Admin Close blocker:** The real modal `modAdmin.Open_Settings`
launcher reuses a hidden default instance after Close, retaining personal/policy
staging. The earlier private-instance close tests did not prove that operator path.
D18's lifecycle clarification requires disposal and fresh saved-state capture.
Focused actual-launcher RED is 3 PASS / 3 FAIL: both activations and Config bytes
pass; hidden-instance retention and both unsaved-choice checks fail. Correct the
real Close handler only after this RED, then preserve the full Settings baseline.
The intermittent Excel process loss at the private-form reopen is separately
recorded; this contract defect does not by itself establish its native cause.
The correction changes the real Close handler from Hide to Unload. Five packages
build/compile with cold-start validation, and focused actual-launcher GREEN is
6/6, including no hidden default instance and both choices discarded. The
combined Settings/Operations regression now passes 187/187; this does not establish
the cause of preceding intermittent automation-process losses or complete user acceptance.

**4be.2 personal preference continuation, 2026-09-13:** Add packaged surface and
real Save/Reload/Reset/Close tests before local persistence. D18's storage
refinement uses current-Windows-user settings with exact invSys-user/warehouse
keys and fixed choices, without Config writes or an Admin capability at the
personal Core boundary. Admin owns its form controls; Operations without Admin,
restart/isolation, invalid-preference fallback and policy/evidence display remain
required gates. Surface RED is 107 PASS / 5 FAIL; real no-write Save RED is
123 PASS / 1 FAIL. All 141 local Settings checks now pass, retaining the
42 profile / 34 policy GREEN checks. The added restart phase encountered an
Excel native fault during extra fixture bootstrap; it now reuses the preserved
second Admin-generated fixture. Keep that failed setup evidence and the two
corrected fixture errors separate from runtime RED/GREEN (see Settings evidence).
The corrected complete run passes 145/145, including a handler-saved choice
restored after Excel restart with no Admin package loaded. Five packages compile,
cold start passes, smoke is 86/86 and the Action Paths capture was inspected.
Static limits and all 105 prior package / 16 protected source pins hold. Operations
personal Settings, Viewer profile/preference consumption and comprehensive new
control activity remain unimplemented; this is not Slice 4be or user acceptance.

**4be.2 Event Detail profile work, 2026-09-13:** Settings keeps General/Event
Tracking and introduces the approved Tracking/Event Detail/Action Paths sections
as nested tabs. D18's schema/editor refinement names the complete display profile,
tables, registered families and required context. Admin owns field selection,
enablement/order and synthetic preview; Core owns profile reads/version saves.
The actual-form surface RED is 72 checks: 64 PASS / 8 FAIL. The staged editor and
real no-write Save action give a clean 97-check RED: 94 PASS / 3 FAIL, including
failure to publish a profile version. The first writer exposed a reload defect:
Excel numeric DisplayOrder requires validated Long conversion before the strict
integer JSON encoder. The corrected candidate completes 107 checks: 105 PASS /
2 FAIL, with all 42 profile checks and every prior GREEN retained. Remaining
failures concern the unfinished personal Settings workflow. All five packages
compile, cold start passes, smoke is 86/86 and the version-2 editor capture was
inspected. Static dynamic-call/duplicate counts and oversized-module limits hold.
Full Receiving passes 854/854, retaining all prior checks and GREENs; all 90
prior package pins and 16 protected source pins match. Read-only package-code
review resolves 20 otherwise unexpected raw hashes as casing-only changes with
identical string literals, alongside the three intended edits/eight additions.
One native Excel build failure left a partial read-fix candidate; preserve it as
failed setup evidence, never a release. Exact reports and the earlier transient
General harness failure are in the Settings evidence record. This work implements
and constrains existing D18 under semantic inheritance; it does not grant new
collection/access rights or complete Viewer, preferences or comprehensive activity.

**2026-09-13, 4be.3 Refresh failure checkpoint:** Implement existing D18
Unavailable/Stale and captured-context rules through actual packaged Viewer
launch, Events, Refresh and Search handlers. Initial RED is 6 PASS / 7 FAIL;
empty-envelope expansion is 14 PASS / 2 FAIL. Final focused GREEN is **16/16**
on `deploy/validation-viewer-refresh-final`. Five packages compile/cold start;
the readable synthetic Stale capture is inspected. Settings retains **187/187**,
populated Viewer passes, packaged smoke is **86/86** and static ratchets hold.
Full chain is **31/31** with ordered live-role child **48/48**. All 140 prior /
10 candidate package pins and 16 protected source pins match; Excel is closed.
Exact scope, failures and final gates are in the maintained
[Refresh evidence](../../invSys_fork/tests/integration/plan022_slice4be_viewer_refresh_results.md).
Only two Operations components change; Core/Domain, canonical schemas and
publication authority are unchanged. The existing D18 contract already requires
this behavior; no architectural exception or new approval is needed. Successful
Shipping supplements still need migration to owning publication boundaries;
complete event groups/detail, source coverage, paging, profiles, recording,
guides/comparison and human UAT remain open.

**2026-09-13, 4be.3 selected-event detail checkpoint:** D18's selected-event/
read-envelope refinement was recorded before implementation. Initial actual
Viewer selection RED is 2 PASS / 5 FAIL; current-state classification expansion
is 29 PASS / 2 FAIL, and resize-after-sign-out expansion is 32 PASS / 2 FAIL.
The corrected `validation-viewer-detail-context` candidate passes **34/34**,
preserving repeated exact keys, multiple entities and unlike UOMs without
aggregation. Core preserves permitted fields from the same projection read;
Operations owns the reused detail surface and profile-version cache. Five
packages compile/cold start; Refresh-failure regression is **16/16**.
Settings retains **187/187**, populated Viewer passes, packaged smoke is
**86/86**, and full chain is **31/31** with its ordered live-role child **48/48**.
Native/default layout and inspected captures pass. Static ratchets hold and
compiled source review finds no unexpected changes.
All 150 prior / 15 candidate package pins and 16 protected source pins match;
Excel is closed. Accepted deployment and NAS remain untouched.
No missing historical identity, publication time or completed outcome is invented.
Raw Note/unknown columns remain excluded and selection performs no source read.
See [Event Detail evidence](../../invSys_fork/tests/integration/plan022_slice4be_event_detail_results.md)
for exact gate scope and earlier setup failures. Complete publication, source
coverage, activity on new controls, recorded conclusions, both Action Path
presentations and human acceptance remain required.

**4be.3 publication and paging D13 entry:** Following code `e66ab6f` / docs
`cd139fb`, protect the remaining D18 published-only read and 100-record page
contracts. Use an Admin-generated disposable snapshot with 5,001 source groups,
out-of-order storage, and repeated-key/unlike-unit lines in the page-boundary
group. Actual Viewer Refresh, Search, selection and Previous/Next handlers must
expose deterministic grouped pages, retain every detail line, and avoid Shipping
authority reads. Architecture records the discovered paging controls before
implementation. This test entry does not accept missing all-source publication,
coverage metadata, the publisher's independent 5,000-group bound or Action Paths.
Expected RED: the current unbounded line list, absent page controls/counts and
Shipping canonical read on successful Refresh. Focused packaged RED on the
unchanged detail candidate is **7 PASS / 9 FAIL**. The fixture loads, filtered
detail retains every exact-key line, and source bytes remain unchanged; paging,
group summaries and exclusion of the Shipping authority-read boundary fail.
The inspected 5,003-line capture confirms unbounded reverse-storage display.
See [group/paging RED evidence](../../invSys_fork/tests/integration/plan022_slice4be_event_groups_results.md).
The publisher's independent 5,000-group selection rule still needs its own
owning-boundary test. No runtime implementation or new GREEN is claimed.

**4be.3 persisted publication test entry:** The next test extends the grouped
Viewer run through the public Admin Generate Inventory Snapshot command.
D18 now names the Core-owned Events JSON artifact, complete-group/source coverage
structure and exact-byte integrity convention. An injected read-only source
fixture supplies the existing snapshot publisher with the 5,001 groups; no
canonical inventory row is fabricated. Require an actual persisted artifact,
verified publication time, exactly 5,000 complete groups, boundary detail and
source-byte preservation. Reader-only clipping cannot satisfy this gate.
Artifact/schema selection implements the approved publication contract under
semantic inheritance. The current package is expected to publish only the
inventory snapshot and fail the separate Events-artifact checks.
The clean combined run is **8 PASS / 19 FAIL**: the 16 grouped-Viewer checks
retain 7/9, Admin publication succeeds, nine Events-artifact assertions fail,
and canonical-byte preservation fails. A separate publication-only diagnostic
is **1 PASS / 10 FAIL** and identifies Inventory as the sole changed source;
Auth/Config/Outbox/operator/source-copy hashes remain unchanged. The existing
Admin command resolves Inventory through a create/schema/save-capable resolver.
D18's source-read constraint governs the next correction: scoped read-only
resolution in Core, with Admin delegating rather than pre-opening via that
writer resolver. The pending Auth provisioning decision is not changed.
Expanded baseline source-scope RED is **1 PASS / 12 FAIL**, proving writable
open, retained source handle and changed Inventory bytes. The scoped Core reader
correction is **4/4 GREEN** for Admin success/read-only/release/byte preservation;
the complete combined run is **11 PASS / 18 FAIL**, retaining grouped 7/9 and
all nine missing-artifact failures. Five packages compile/cold start; compiled
comparison shows only Core snapshot/Admin delegation and the new Core class.
Detail **34/34**, Refresh failure **16/16**, Settings **187/187**, populated
Viewer and smoke **86/86** pass; layout/capture and static ratchets hold.
Full-chain recovery exposed missing imports in the generated Create Warehouse
source harness and an LF-only class export that direct VBA import treated as
code. The 16 explicit source-harness dependency lists and class export format
are corrected; standalone Create Warehouse is **15/15**. The packaged candidate
is unchanged. These were harness failures, not behavioral RED or a new contract.
A later native Excel fault remains recorded with unresolved cause. The unchanged
retry passes full chain **31/31**, including completed ordered live-role **48/48**.
Full Receiving passes **854/854**, retaining every prior check identity and GREEN.
All 165 prior + five candidate package pins and 16 protected source pins match;
Excel is closed, static/schema checks hold, and all 19 changed scripts parse.
See [publication/source-read evidence](../../invSys_fork/tests/integration/plan022_slice4be_publication_results.md)
for gate scope, earlier failures and final Receiving/preservation evidence.
This source-read correction does not create the Events artifact or implement
paging, all-source coverage or Action Paths.

**4be.3 next publication D13 scope:** Extend the actual Admin snapshot test with
an activity pair generated by the real Settings Save Value handler. Require one
exact ActivityId group containing both original RecordIds and the actual result,
earliest-record chronology, package/policy provenance, named source coverage and
reconciled Inventory/Activity counts after the global 5,000-group bound. The
normative wire detail separates ShippingBOM from station-local ShippingHolds;
neither missing owner coverage nor local holds imply complete warehouse history.
Current tests counting coverage entries alone are insufficient. Runtime publisher
implementation and actual Designs/Shipping source-population proof remain pending.
The publication-only diagnostic now records **6 PASS / 15 FAIL** through the
actual Admin handlers: the correlated pair and all four source-read checks pass;
the missing artifact, group/provenance/coverage assertions fail. Two combined
attempts stopped earlier in Viewer setup and are not publication RED or accepted
regressions. The diagnostic route still cannot claim acceptance GREEN. All 170
package pins and 16 protected source pins match; Excel is closed. See
[mixed-source publication RED](../../invSys_fork/tests/integration/plan022_slice4be_publication_coverage_red.md)
for exact failures, the fixture metadata correction and outstanding gates.
Code `68ecec1` adds Designs' `PUBLICATION_EVENTS` query through the existing
Core/Domain dispatcher; candidate `deploy/validation-designs-publication-source`.
D18 fixes its primitive envelope and read-only, captured-target, clean-borrowed-
source, header and redaction rules. Owner-applied lifecycle events in an Admin-
generated fixture prove **3 PASS / 12 FAIL RED -> 15/15 GREEN**. The combined
publication-only diagnostic moves from 9 PASS / 27 FAIL to 21 PASS / 15 FAIL;
the remaining failures concern the missing Events artifact. All five candidate
packages compile with cold-start references checked. Detail 34/34, Refresh 16/16,
Settings 187/187, populated Viewer, smoke 86/86 and chain 31/31 with ordered
live-role 48/48 pass. The full Receiving retry retains every prior 854/854 GREEN
identity after a recorded 606-pass native-focus setup interruption. Static
dynamic-call/duplicate/oversized limits hold; 175 package pins and 16 protected
source pins match, with Excel closed. This is a
prerequisite to actual all-source publication, not completion of the Admin
publisher or authorization of a Viewer authority fallback. See
[Designs source evidence](../../invSys_fork/tests/integration/plan022_slice4be_designs_publication_source_results.md).

**4be.3 Shipping publication test extension:** D18 now names the permitted
Shipping current-state line fields and explicit Warehouse/Station profile scope.
Prepare two BOM components and a held shipment through actual packaged Box
Designer, Box Maker, Add and Hold handlers in an Admin-generated seeded fixture.
Extend the public Admin publication test to require every exact component/package
key and held-line reference, truthful current-state classification and coverage,
and unchanged canonical/local-source bytes. Preserve the controlled 5,001-group
source before the owning fixture actions publish their inventory snapshots.
Reconcile Activity counts with the actual records those handlers create. Expected
RED remains the missing Events artifact; fixture or harness failure is not RED.
No runtime implementation or comprehensive publication GREEN is claimed here.
The expanded unchanged-package diagnostic is now **26 PASS / 21 FAIL**, with no
harness failure. All 36 previous check identities and 21 GREEN are preserved;
the four real-owner fixture checks and local-source preservation pass. Six new
Shipping artifact assertions fail alongside the previous fifteen missing-artifact
checks. Four earlier harness failures are retained separately. The held-line
capture is inspected and still shows pending sync, not a completed shipment.
All 175 package pins and 16 protected source pins match with Excel closed. See
[Shipping publication RED](../../invSys_fork/tests/integration/plan022_slice4be_shipping_publication_red.md).
The Core-owned publisher, declared Shipping owner read, combined Viewer gates
and complete Release 1 acceptance remain required.
Implementation now declares the fixed Core-to-Operations Shipping owner read in
D18. One additional literal Application.Run site is a scoped D12 bridge exception
to the current eight-literal maintenance baseline; no arbitrary target dispatch
or new same-project dynamic call is authorized. Core retains the Inventory
Boolean/path contract and supplies a separate Events notice, reported by the
explicit Admin snapshot action. Bootstrap without a matching allowed target
defers Events without retargeting. The existing 26/21 packaged RED protects the
first implementation; atomic failure and unavailable-source assertions must also
be exercised before publication acceptance. No GREEN is claimed by this entry.
The snapshot orchestrator has a scoped six-line size exception (1,753 to 1,759)
for the optional Events result, per-call publisher instance, source capture and
publication delegation, including a local result string before copying the notice
to the caller's optional output. The implementation stays in the new bounded Core class;
this exception does not approve other oversized-module growth. Shipping reuses
its existing field-escape routine through a direct same-project call, preserving
the duplicate-body baseline rather than adding another codec implementation.

**2026-09-13, Events publisher candidate:** Calibrated public-Admin RED is
27 PASS / 39 FAIL; `validation-events-publication-atomic` passes all66 focused
checks for complete groups, exact permitted owner lines, source bytes, truthful
coverage and failure/recovery. The Shipping fixture now adds its unknown column
through a calibrated VBA edit and requires a read-only reopen. Earlier unreadable
PowerShell-COM fixture output is not evidence for a runtime Shipping repair.
All five packages compile; Detail34, Refresh16, Settings187, populated Viewer and
smoke86 are GREEN. Two unchanged atomic full-chain attempts fail natively at
projection rebuild (live32/1); an unsaved compile-invalidation diagnostic passes
live48/48 but is not acceptance. A separate saved-compilation candidate preserves
all199 component hashes yet fails natively during Production (live27/1); saving
compilation state does not resolve the gate. Full-chain acceptance remains pending.
Receiving passes854/854 with every prior identity/GREEN retained and no duplicates.
The pre-publication baseline passes the same full chain31/31 with live48/48;
candidate-specific failure remains unresolved. A separate local-output-string
experiment changes only the snapshot owner and passes full chain31/31 with live48/48.
The source now uses this internal result handling. The fresh reviewed candidate
compiles all five packages and passes chain31/31, live48/48, publication66,
Detail34 and Refresh16. Settings stops54/1 on reopen after Close. The optional
unsaved constructor trace reproduces that failure during layout construction;
its result is diagnostic, not behavioral RED or acceptance GREEN.
Finer unsaved layout tracing completes187/187. The subsequent uninstrumented
Settings regression also passes187/187 with every check identity/GREEN retained
and no duplicates, closing that gate without identifying a native-crash cause.
The experiment alone does not establish a general native-crash cause or repair.
This preserves D18's separate per-call results and introduces no shared last-result
slot, new operator behavior or architectural contract. The fresh build must retain
focused66 and every required regression before the publisher implementation is committed.
The added actual Box Maker assertion proves immediate publication of its
durable owner event before explicit Admin publication; all prior65 checks remain
GREEN with no missing or duplicate IDs. Preservation matches
175 historical and15 candidate package pins,15 exact protected sources and one
reviewed visibility-only Shipping source change. Three packaged additions/six intended edits and34 capitalization-
only changes with identical literals are reviewed; no unexpected difference.
Current reviewed static counts are206 components/5700 procedures/127035 lines,9 literal/45 unresolved
calls,189 duplicate groups and28 size ratchets with only the stated six-line
exception. This advances publication; combined Viewer projection/paging, complete
Operations/Admin coverage, Action Paths and physical/human acceptance remain open.
Native fault processes created after the passing live runs also require phase
attribution; passing chain rows alone do not prove clean native shutdown. Do not
infer an Office installation defect or accept instrumentation as the final gate.
The reviewed candidate now also passes populated Viewer, smoke86, native Shipping
layout1/1 and reusable Production/restart2/2. Receiving stops690/1 at a rejected
native navigation readback; the required854-check gate remains pending. Excel
closes after the failed run. Isolated final-chain phase tracing passes13/13 and
observes swallowed first-package Close exceptions, but does not reproduce the
native fault or establish its cause. These diagnostics do not replace acceptance.
The unchanged focused Receiving navigation diagnostic then passes266/266, including
the previously interrupted mouse transition. A fresh full854 regression is running;
the focused result is not a substitute for its complete acceptance evidence.
The final-phase cleanup trace also exposes Core being closed before dependent
packages. A bounded D13 harness test executes the actual finally blocks:5/8 RED
becomes13/13 GREEN after reversing owned-workbook cleanup order, with no implicit
save. Ordered-child exit/report decisions remain4/4. No XLAM or architectural
contract changes; real-Excel after-tracing and full-chain verification are pending.
Real-Excel after-tracing now passes13/13 with zero package-close exceptions and
Excel closed. The full Receiving retry passes854/854 with all preceding identities/
GREEN retained. A fresh full chain nevertheless fails earlier at projection rebuild
(top4/1, live32/1), before either corrected cleanup phase; no crash repair is claimed.
The subsequent retry finishes31/31 and live48/48, but Windows records a native
Excel fault just after the report. Its verified empty recovery child closes without
saving; all three loaded add-in files remain unchanged. Clean native shutdown and
publisher acceptance remain open. Both released-deleted-handle and unchanged-handle
live diagnostics pass48/48 with the same check set, so the maintained projection
fixture remains unchanged. Independent published-only Viewer tests can advance
under existing D18 while this acceptance limitation stays explicit.
The new actual Settings -> Admin publication -> Viewer test completes **9 PASS /
14 FAIL**,23 unique checks and no harness failure. Activity/detail and publication
time are absent; Viewer still enters the Shipping supplement reader and does not
reject damaged, incompatible, wrong-warehouse or missing Events artifacts. Current
visibility Save handlers do not republish, but the positive visibility restore
case fails. Context invalidation and file preservation pass. Optional foreground
captures remain setup failures, not acceptance. See
[published read RED](../../invSys_fork/tests/integration/plan022_slice4be_published_read_results.md).
The reader implementation uses D18's `EVENTS1` primitive wire refinement: current
policy and publication metadata accompany every allowed contributing line, with
named detail fields following the eighteen compatibility slots. Core owns the
read and visibility decision; Operations parses the projection without Shipping
authority fallback. Existing legacy serialized envelopes retain limited coverage.
This is an isolated candidate until focused, package, regression and visible gates
pass; paging and complete Events/Action Path acceptance remain independently open.
The initial reader candidate now passes **23/23** focused checks with all RED
identities retained, all-five-package compile/cold-start, Refresh16/16 and Detail
34/34. Detail fixture calibration publishes the same synthetic lines through Core
and preserves their source bytes and existing assertions. Full group/paging,
Shipping presentation, maintenance, role/chain and visible gates remain open;
the candidate is not a completed runtime slice.
The volume/group suite advances from7/9 to8/8 with all16 identities and every
prior GREEN retained. Published-only read is now GREEN; paging, page counts and
unlike-unit summaries remain RED. Static evidence is207 components/5709 procedures/
127173 lines,9 literal/45 unresolved calls,189 duplicate groups and28 size ratchets.
Four added maintenance candidates, including obsolete reader helpers, require
review; no new size, duplicate or dynamic-call exception is approved.
The next isolated Operations candidate implements the existing D18 group/page
contract in `cEventPageProjection` and the Viewer form: cached loaded groups,
100 matching summaries, Previous/Next and count controls, with full source-line
selection and explicit Multiple values. The existing8/8 RED protects it. The
same package compiles and an expanded Day/All test establishes24PASS/1FAIL for
verified UTC events compared against local Now. UTC-aware comparison follows
that focused RED; no architecture or timestamp-provenance weakening is adopted.
The UTC candidate compiles all five packages and passes group/page16/16, reader
25/25, Refresh16/16 and Detail34/34. Expanded real-form geometry records25/4 RED:
the status label exceeds the client height by1.4 points at each tested size.
Moving it up four points preserves its height and passes all four resize checks;
foreground and human evidence remain separate. Mixed-source equal-time ordering
then establishes29/1 RED with all25 prior GREEN retained. The corrected pager reads
the named SourceKind field and follows D18/the publisher's SourceKind/SourceId/
Source tie order. No architectural rule changes. The new isolated candidate
compiles all five packages and passes cold start, reader30/30, group/page16/16,
stale Refresh16/16 and Detail34/34. Exact prior regression identities remain;
static call/duplicate/size ratchets hold and protected package/source pins match
with Excel closed. Shipping presentation, remaining filters, full role/chain and
visible acceptance stay open; runtime remains an uncommitted candidate. Excluded
declaration-placement and bootstrap setup
failures are retained in the linked paging evidence; neither is behavioral RED.
See [package cleanup evidence](../../invSys_fork/tests/integration/plan022_slice4be_package_cleanup_results.md).
See [Events publication evidence](../../invSys_fork/tests/integration/plan022_slice4be_events_publication_results.md).

**Shipping current-state presentation:** D18's explicit state-presentation mapping
preserves one package/alternative summary with every component detail, using the
existing named BomId/BomVersion fields as a presentation association and keeping
SourceId unavailable. The protecting `-CheckViewerShippingState` run retains all30
reader checks and prepares state through real Box Designer, Box Maker, Add/Hold
and Admin publication handlers. Complete RED is50PASS/6FAIL: package summaries,
component detail grouping/search and the held-line label fail; owner publication,
identity/provenance and read-only preservation pass. Correct Core's published
display mapping and the Operations pager/detail selection against these failures.
No current-state association may become a fabricated historical event identity.
The corrected isolated candidate compiles all five packages and passes56/56,
retaining the complete RED check set, plus paging16/16, Refresh16/16 and Detail34/34.
Static call/duplicate/size ratchets hold. Full-chain31/31, live48/48 and source
integration15/15 pass, but Windows records another native Excel combase/c0000005
fault during the run. Its cause and clean native execution remain unproven.
Excel is closed, package/source preservation checks pass, and prior generated
reports were restored. Runtime remains uncommitted pending broader acceptance
and maintenance. See [Shipping Viewer evidence](../../invSys_fork/tests/integration/plan022_slice4be_shipping_viewer_results.md).

**Next D13 filter gate:** Implement D18's explicit View/family/source/recorded-outcome
selectors through actual ComboBox change handlers. Protect AND-combined matching
across complete loaded groups, internal reservation labels, unavailable metadata,
page-one reset, full detail, context invalidation and responsive Events-only layout.
Include a staged date change: selector changes retain the last refreshed range;
only successful explicit Refresh applies the new range.
The fixture uses a declared synthetic wire only after the real publication/reader
checks pass. No filter may read or publish, infer a sequence conclusion, or reveal
policy-hidden metadata in its choices. Establish behavioral RED before adding
the controls to the Operations candidate.
The complete packaged baseline is33PASS/22FAIL with all30 reader GREEN retained.
The initial implementation compiles five packages and passes55/55; duplicate-body
groups increase189->190 from four identical event wrappers. Replace those with
one typed Operations binding that disconnects on form close and reverify against
the same behavior tests. This is maintenance within the approved filter contract,
not an exception to the duplicate ratchet. Selector activity coverage remains
explicitly pending in4be.1; population/rendering is never a user action.

The revised binding candidate compiles all five packages, passes Operations cold
start and the complete55/55 filter/reader checks. Maintenance returns to189 duplicate
groups, with9 literal/45 unresolved calls and28 size ratchets unchanged. Foreground
capture stops after53 behavior/layout checks because the Viewer is not foreground;
visible acceptance remains open. Shipping regression stops after30 reader checks
with a COM failure during the actual Box fixture action. Its remaining empty Excel
process is verified as a child of the failed test and eventually closes after normal
Quit; no forced termination bypasses the retained-content guard. The unchanged retry
passes Shipping56/56; paging16/16, Refresh16/16 and Detail34/34 also pass, retaining
every prior check identity. Visible capture remains unavailable: current read-only
observation finds no foreground window or readable input desktop. See
[filter evidence](../../invSys_fork/tests/integration/plan022_slice4be_event_filters_results.md).

**Inventory layout preservation entry:** Source review finds that the Events-only
paging gap also reduces the Inventory list by40 points while its controls are
hidden. D18's refinement preserves the accepted Inventory list area. Add a real
tab/geometry check to `Slice4beViewerFilters.ps1` expecting the prior12-point gap
above the bottom Close control. Establish RED before correcting Operations
`frmInventoryViewer.ConfigureEventFilterGeometry`; all prior55 filter checks and
the supported resize sizes remain required.
The expanded run records55PASS/4FAIL, confined to the four new Inventory-size
assertions. Correct the Inventory list's bottom anchor to the existing Close
spacing; preserve Events paging geometry. The isolated correction requires fresh
build/compile, expanded59-check GREEN and regenerated maintenance evidence.
The corrected package compiles all five XLAMs and passes59/59, retaining all55
earlier checks. Individual maintenance comparison then finds a missed violation:
`modWarehouseSync` has1759 lines versus committed1753. The stable count of28
oversized modules was insufficient evidence; earlier size-ratchet claims are
superseded and no exception is approved. Review the private, unreferenced
`AppendLocationSummariesSync` before any removal, then require compile and public
snapshot/publication regressions. Keep the running full-chain source unchanged.

The Inventory-layout chain finishes31/31, live48/48 and source integration15/15,
preserving package bytes and prior reports, but records an Excel unknown/c0000409
fault. After termination, explicit body/reachability review removes the unused
location-summary routine and six obsolete Viewer helpers plus their unused table
constant. Core WarehouseSync shrinks to1727 against1753; ViewerData becomes184.
The old Slice4w static assertions are reconciled to current typed handlers and D18
publication reads, retaining all12 identities:8/4 stale assertions become12/12,
and12/12 remains after removal. This is not behavioral RED. The new maintenance
candidate compiles all five packages and cold-starts Operations; its publication,
expanded filter, chain and individual maintenance checks are in progress.
Publication now passes82/82, retaining all66 earlier publication identities plus16
grouped Viewer checks. Compiled comparison covers202 components and isolates the
cleanup to the two reviewed Core modules. Fresh maintenance checks every one of28
oversized modules against its committed limit: none grows; all six new modules and
59 new procedures fit1000/200-line limits. Counts are209 components/5727 procedures,
1136 scanner/1138 reviewed candidates,189 duplicates and45 unresolved/9 literal calls.
The size violation is corrected without an exception. Final filters pass59/59;
the final chain exits0 at31/31, with live48/48 and source integration15/15.
Package hashes and original reports are preserved, but Windows records another
combase/c0000005 Excel fault. Source preservation passes and Excel is closed.
This publication/Viewer source checkpoint changes no accepted deployment or NAS
workbook. It does not complete Slice4be or native/visible acceptance.
Catalog version8 still registers31 controls, including only
one Production and one Admin control; comprehensive coverage/recording/comparison
remain required beyond this publication and Viewer checkpoint.

**Source checkpoint:** Code `1d6a2ad` commits the publication/Viewer implementation,
restored Inventory space, reviewed cleanup, protecting tests and source-harness
dependencies. Earlier notes that runtime remained uncommitted describe preceding
candidates. This commit is not a deployment or complete Slice4be/Release1 acceptance.

| Subslice | Public packaged test and meaningful RED | GREEN outcome |
|---|---|---|
| 4be.1 Coverage and activity foundation | Map every reachable Operations/Admin control to its actual callback/handler and completion source. Run eligible Operations and Admin actions; expect missing stable activity IDs, source/outcome correlation and publication coverage. Test programmatic invocation does not impersonate user input. | Shared non-authoritative NAS Activity store, exact source references and actual outcomes; complete catalog with explicit exclusions; Admin station history labelled unavailable; required audits and business schemas unchanged. |
| 4be.2 Settings policy and profiles | Through actual Settings open/save/reset handlers, expect missing Event Tracking tab, policy/profile persistence, denials and personal view preference. Reject cross-target/stale-version saves and unknown controls/fields. | Admin General/Event Tracking tabs; D5 Core policy/profile commands; required-versus-optional tracking; safe defaults, atomic compatibility settings and unknown-column preservation. Operations personal Settings works without Admin installed; invSys-user/warehouse preference isolation and restart. |
| 4be.3 Comprehensive Events | Through Viewer launch/Refresh/filter/selection, expect missing family/source/outcome coverage, identities/detail lines and bounded paging. The published-read test uses actual Settings activity and Admin publication, then exercises integrity/schema/warehouse rejection, current visibility changes, stale recovery and context invalidation through Viewer handlers. | 5,000 complete published groups, 100-row pages, all contributing lines, labels/zone/freshness/coverage, policy-aware display and read-only authority. |
| 4be.4 Recorded sequence and conclusions | Start Recording, use actual Receiving/Production/Shipping/Boxing/Admin handlers, Stop/Cancel and evaluate. Expect absent ordered controls, cross-form correlation and terminal results; accepted-but-unapplied, missing or interrupted evidence must not pass. | One explicit actor/warehouse sequence across permitted forms/submissions, immutable attempts/results, 256-action and 1 MiB limits, confirmed-versus-pending/failed/cancelled/incomplete conclusions; policy changes, tracking failure, retries and source authority protected. |
| 4be.5 How-To and comparison | Select events or a captured sequence, edit/save/search a guide, choose How-To/Diagnostic/Compare both and exercise export/import. Expect missing guide/evidence pairing, method preference and distinction between authored steps and observed results. | Both presentations of one record; stable selection on switch; capability-gated immutable guide versions; validated export/import with origin-only evidence; older-release warning and current policy respected. |
| 4be.6 Release and visible comparison | Run protecting regressions after focused GREEN; do not create artificial RED for already accepted behavior. | Full five-package build/compile/initialization, layout/static/live-role/full-chain/restart proof and visible user comparison before Slice 4be acceptance. |

**4be.4 late-result isolation test entry:** Following source00a7ab6, extend the
packaged recording gate with `-CheckRecordingIsolation`. Real Viewer Start/Stop
and Admin Save Value prepare two distinct runs. Replay only the prior owner's
observed result before/after actual policy changes; expect the global
FinishAction interruption to close the newer run incorrectly. D18 requires
same-policy idempotence, old-policy rejection without cross-run interruption,
unchanged journals/activity/config and normal current-run Stop. Preserve all117
prior checks. This enforces existing sequence ownership; no new architecture
or business action is introduced. The frozen candidate now records127PASS/3FAIL
across130 unique checks, retaining all117 prior GREEN identities with no harness
failure. Same-policy replay passes; changed-policy replay rejects the old result
but closes the newer journal, disabling active status and normal Stop. Activity
and config bytes remain unchanged, all five package hashes are preserved and
Excel closes. Next scope FinishAction interruption to its owning sequence,
preserving policy rejection and current-run failure handling, then prove GREEN
and applicable release gates. No runtime correction or isolation GREEN is claimed.

The follow-up matrix also protects FinishAction's append-failure and exception
paths. A conflicting disposable outcome file and a calibrated exception after
action/context resolution exercise those actual branches, using only previously
observed owning outcomes. Restore fixture bytes/exception selection afterward. Repeat the two
failures for the active sequence as positive controls: an owning failure must
still append Incomplete Close. Preserve every original journal entry and activity/
config byte. The first extension stops75/1 on an invalid repeated-value fixture
(the owner correctly reports UNCHANGED); this is a harness failure, not RED.
Use distinct actual Settings values in the corrected run before implementation.
The corrected-value run is94/11: six valid policy/append failures, five invalid
exception-fixture expectations. A CVErr coercion did not establish the intended
error branch; replace it with explicit test-only Err.Raise in the unsaved project
and require the production handler's sanitized response. Do not claim all11 as RED.
An exact-case instrumentation attempt stops68/1 before the matrix; a read-only
package probe calibrates VBE's identifier casing. The final calibrated matrix
records96/9 across105 checks: stale policy, append and exception paths each fail
the same three isolation checks; both owning-failure controls pass. No harness
failure remains in that run. Implemented internal InterruptAction checks exact
Context/SequenceId at FinishAction's three failure sites. Five new candidate
packages compile; static candidate IDs, duplicates/dynamic calls and all28
individual oversized-file limits are preserved. The154-check focused gate and
full chain remain required. The first combined candidate run reaches109/1: all37
isolation checks pass, then reader entry loses RPC. The exact owned empty child
exits after normal Quit; the forced-stop guard refuses termination while its
window remains. A later process check confirms closure and five candidate hashes
remain unchanged. Retry uses unchanged packages/tests; no154-check or full-chain
GREEN is claimed yet. This is an implementation of existing D18, not a new contract.
The unchanged retry subsequently passes154/154, exit0, retaining all117/130/105
prior check identities. Compiled hashes differ only for the two intended Core
modules across211 components. The same candidate completes full chain31/31,
live-role48/48 and Create Warehouse15/15, exit0. Its five package hashes and three
restored tracked reports remain unchanged; no Excel Application Error1000 was
observed in the chain window. The prior RPC/recovery failure remains unresolved.

The passing chain leaves one residual Excel instance. Initial read-only inspection
found zero workbooks and three add-in projects. Normal Quit preserved all three
add-in file hashes, but the process persists with a visible native dialog whose
text is unavailable through automation. No forced termination occurs. Registration
restoration is a suspected cause, not proven provenance; null shutdown COM values
do not prove an empty process. Excel closure and native/visible acceptance remain
open. Do not start another Excel job or build while it remains. A read-only
snapshot verifies260 package pins,15 protected source files and the existing
reviewed Shipping visibility-only difference, explicitly recording ExcelClosed=False.
This checkpoint implements existing D18 ownership without changing architecture.
Next prove an actual unclosed recording across a fresh Excel process, then real
Operations/multi-event and deferred-result sequences, conclusions, guides and both
presentations. Full Slice4be and Release1 acceptance remain open.

**4be.4 actual cold-interruption test prepared:** `-CheckRecordingRestart` retains
the packaged reader baseline and adds12 checks. Start plus a real Admin Save Value
must persist an unclosed three-entry journal. The test verifies the exact initial
Excel HWND/process, candidate project paths and disposable saved workbooks before
terminating that held process; it never deletes a Close record or uses a runtime
reset to simulate interruption. Fresh Excel must display Interrupted through the
actual library controls, preserve fixture bytes and package hashes, leave ordinary
actions unsequenced, and give a new explicit run a distinct identity. Existing
limits/storage/isolation gates run separately and remain required. Three changed
PowerShell parsers and whitespace checks pass. Runtime execution and fresh-driver
compile remain unverified while the prior chain's visible Excel dialog remains.
No RED/GREEN or contract change is claimed; a setup/ownership/compile failure must
not be counted as behavioral RED.

**2026-09-14 execution update:** MSAA identifies the preceding chain residual's
Document Recovery prompt. The exact process is reverified; selecting and verifying
**Yes, I want to view these files later**, then OK, permits normal exit without
force or recovery-file deletion. The first cold-restart gate completes111PASS/
1harness failure: all108 earlier checks and real durable interruption/distinct
Excel process checks pass. Fresh-process instrumented Viewer opening produces
the compiler error **Error accessing file. Network connection may have been lost.**
The code pane selects cOperationsAnchorManager line75; cause and actual network
loss are not established. Acknowledging the error and resetting only that failed
test from verified break mode lets the harness exit and Excel close normally.
The five candidate hashes remain unchanged. This is not product RED or full
restart GREEN. The next controlled run adds pristine Viewer launch before fresh
driver installation, retaining all12 original restart checks; no runtime behavior
or D18 contract is changed. Earlier RPC/native reliability remains unresolved.

The calibrated rerun also records111PASS/1harness failure. Its pristine Viewer
launch, through the existing packaged action wrapper, encounters the same compiler
error before any fresh-process driver code is installed. Editing those drivers
is not a necessary trigger. Keep the scope precise: same PowerShell controller,
intentional prior Excel interruption, then a fresh Excel process. Ordinary cold
operator startup has not been proven defective by this comparison. After error
acknowledgement, COM debugger access is unusable; the exact failed test process
is terminated to release the waiting harness. Its subsequent RPC-unavailable
report is a cleanup consequence, not another independent crash. Both runs are
terminal, Excel is closed, and260 package pins plus15 protected sources and the
reviewed Shipping visibility-only change pass preservation. No source/package
runtime change is made. Next separate the reader's PowerShell controller, pass
fixture input only in memory, and calibrate pristine launch there before adding
drivers. Cold-restart and full native/visible acceptance remain open.

**2026-09-14 separate-controller comparison:** Extracted14 existing recording
fixture helper bodies unchanged and retained all13 restart check identities.
A fresh-reader worker receives fixture data through private stdin, never command
arguments/files/logs, and emits only fixed stages and boolean check records.
Protocol/redaction checks pass6/6 without creating Excel or authority. The full
run records112PASS/2harness flags for one compiler failure: all108 earlier checks,
the interruption/process checks and a new distinct-controller check pass, but
pristine Viewer launch still fails before fresh-driver installation. The original
creator controller remains alive. After error acknowledgement fails to release
the invocation, only the exact failed worker Excel process is terminated; both
controllers then exit and Excel is closed. No runtime package or contract changes.

A standalone pristine-startup control uses Admin Generate Warehouse/auth fixtures
and the existing Viewer action wrapper, without VBA edits or an interruption in
that testcase. It passes6/6, including first launch, form reuse and fixture/package
preservation, and Excel closes normally. This rules out universal pristine-launch
failure in that setup; it does not establish cause because the control also runs
bootstrap in the fresh Excel process. Next use a neutral coordinator: the creator
fully exits before the reader opens the same saved fixture, with no repeated
bootstrap or driver edits before pristine launch. Keep the private fixture transfer,
registry restoration and cleanup ownership explicit. This is continued D13
calibration under existing D18, not a replacement for interrupted-reader, full-role
sequence, conclusion, guide/comparison or native/visible acceptance.
Final closed-Excel preservation verifies260 package pins,15 protected source files
and the existing reviewed Shipping visibility-only change. Six changed PowerShell
files parse and83 local document links resolve. Runtime/static/package source is
unchanged; no new full-chain or human-acceptance result is claimed here.

**2026-09-14 creator-exit recovery proof:** Code checkpoint `2e6b09a` adds a neutral
coordinator that never opens Excel. The creator runs the108-check foundation,
starts a real Viewer recording and records an actual Admin Save Value, interrupts
only its verified disposable Excel, transfers the same saved fixture privately,
and exits. The fresh reader verifies creator exit and launches pristine Viewer
before driver installation, without repeated bootstrap. The complete run passes
**123/123**, exit0, with normal final Excel closure and all112 prior passing check
identities retained. It shows the original unclosed journal as Interrupted, never
resumes or concludes it, retains the original attempt/outcome, preserves read-only
fixture bytes, keeps ordinary work outside it and gives a new explicit recording
an independent identity. All15 restart checks pass on the unchanged runtime.

The current-user-only private transfer passes5/5 after correcting implicit console
decoding with explicit UTF-8 streams; worker protocol/redaction passes6/6. An initial
coordinator path-argument failure precedes Excel/checks and is corrected. These are
harness setup results, not D13 product RED. Cleanup restores settings/removes only
the disposable fixture after both controllers and Excel exit; a live process
explicitly defers cleanup. No fixture payload enters command arguments or reports.

Final verification preserves260 package pins,15 protected sources and the existing
reviewed Shipping visibility-only change; seven changed scripts parse. Runtime
remains `99a69aa` / `deploy/validation-recording-isolation`. No new architecture,
runtime, package, static-baseline, full-chain or human-acceptance change is claimed.
This closes the focused interrupted-reader gate in this lifecycle; it does not
explain every prior native/compiler failure. Under existing D18/D13, next prove
one recording spanning real Operations/Admin actions and multiple submissions,
including exact multi-event references and deferred published owner outcomes.
Conclusion evaluation, How-To/Diagnostic/Compare, versioned guides, comprehensive
coverage and visible NAS/operator acceptance remain required.

**2026-09-14 Operations sequence/evaluation RED:** Code `b2c995f` adds
`-CheckRecordingOperations`, retaining the108-check reader foundation. The actual
Receiving Ribbon callback reuses its captured form/workbook while an unrelated
workbook is active. Actual Add/Confirm and Admin Save Value produce nine actions,
18 observations and20 immutable journal entries across two submissions. With
automatic application withheld, the first two source references remain among the
second submission's four; the library displays all six observations of references.
The later owning Admin processor applies four distinct events. Published groups
move from zero to four and preserve every exact System_Key. No generated identity
or canonical outcome is invented by the recorder.

The corrected final run reports **135PASS/2FAIL**, with normal Excel closure.
Only the pending/applied library's missing Evaluate action fails; every108
foundation identity and all130 prior passing identities remain GREEN. Journal
bytes, unknown staging columns and unrelated workbook bytes/content are preserved.
Saved=True is not a substitute for those data checks: historical Receiving GREEN
already records a changed Saved flag, and the new stage diagnostics preserve the
file/content throughout. The original journal pin check permits separate derived
evaluation records, as D18 requires; it never authorizes rewriting observations.

Calibration failures remain explicit in the maintained evidence: an Add fixture
needs distinct references/actual Refresh readiness; a capture-enabled run fails
the existing foreground guard19PASS/1harness failure; an earlier corrected run
fails at publication120PASS/1harness failure with RPC unavailable. Its remaining
empty Excel instance closes normally and preserves its three loaded add-in files.
The later135/2 run does not explain the native failure or establish visible UAT.
Three changed scripts parse. Runtime remains `99a69aa` in
`deploy/validation-recording-isolation`; no runtime, package or architecture
change is made and no new full-build/static/chain/human gate is claimed.
Final closed-Excel verification preserves260 package pins,15 protected sources
and the existing reviewed Shipping visibility-only change;83 local document
links resolve. The ignored aggregate is `recording-operations-final-preservation.json`.

**2026-09-14 expectation/evaluation implementation refinement:** The normative
D18 subsection **4be.4 expectation and evaluation refinement** now specifies the
shared ordered expectation editor, optional captured expectation and separate
evaluation draft. This is semantic inheritance of the approved D18 rules, not
a new architecture approval or evidence of implemented behavior.

Operations owns `frmActionPathExpectation`, the Viewer Expected conclusion entry,
and library Expected conclusion/Evaluate/result surfaces. Core validates and stages
the exact definition, freezes it only on recording Close, evaluates the selected
run against the explicitly loaded Events evidence, and appends immutable results.
New journals use schema 2; complete schema-1 journals remain readable without
rewriting. Evaluation files use the fixed Evaluations child directory and never
replace journals. Training editor/evaluation controls are catalogued exclusions
from activity capture; saved training records carry their own provenance.

The test-first sequence extends the actual Operations/Admin fixture through the
real editor and Evaluate handlers. Protect None/default and cancelled draft,
stable ordered StepIds and distinct repeated occurrences, retry behavior, command
completion versus all terminal source events applied, partial application, explicit
Refresh, stale/restricted/corrupt evidence, current context/policy, and immutable
result provenance/read permissions. Expected initial RED is absent expectation
controls and evaluation results on the unchanged compiled candidate; harness or
compile failure does not count. Preserve every prior GREEN identity and package.
Only after focused behavioral RED may the matching Core/Operations runtime change.

The existing presence RED must not be satisfied with a button that performs no
evaluation. Schema compatibility, layout, five-package compile, maintenance,
live-role/full-chain regressions and visible evidence remain required. Full role
coverage, How-To/Diagnostic/Compare, versioned guide authoring/import/export and
NAS/operator acceptance remain within the unchanged goal.

**Expectation editor/partial-application RED (2026-09-14; code `564a90e`):** The new
`-CheckRecordingEvaluation` gate uses the actual library/editor controls and a
fixed four-step expectation spanning Admin Save and two Receiving submissions.
Admin's real BatchSize save limits application to three events, then the next
processor call brings the total to four. Pending, partial and applied publications
are separate verified fixtures. The unchanged candidate ends **153 PASS / 33 FAIL**,
with no harness exception and normal Excel closure, retaining all 135 previous
passing identities and all 108 foundation checks. Expected failures are missing
expectation/evaluation controls and saved results. Preservation checks do not
substitute for the failing save/append assertions.

An earlier run ends 138/23, including one RPC harness failure at the second
processor call; its three-event partial fixture passes. The verified empty child
Excel instance closes normally with its three add-in files unchanged. The later
complete fixture does not resolve that native failure. Final verification retains
260 package pins, 15 protected sources and the reviewed Shipping visibility-only
change. Runtime/packages remain unchanged. Extend the remaining expectation,
matching, policy, stale evidence and result-integrity tests listed above before
implementing those behaviors; full Slice 4be/Release 1 acceptance remains open.

**Expectation/evidence semantics RED (code `1746f88`):** `-CheckEvaluationContracts`
extends the preserved 153/33 gate with real form-driven captured expectation and
rejected-then-successful Admin actions. It protects retry matching, distinct repeated
steps, ordered missing steps, command-only wording, cancellation, ordinary Viewer
evaluation, current restrictions and historical eligible capture. Evaluate before
Refresh must keep the older loaded publication; failed Refresh must make that
evidence stale until another successful explicit Refresh. These are tests of the
existing D18 refinement, with no architecture or runtime change. Missing controls
and conclusions are expected RED; fixture/VBA/COM failures remain separate.

The first run reports 160 PASS / 55 FAIL with normal closure. A sign-out assertion
is then tightened to inspect result text even when hidden, require a previously
staged expectation, and verify no derived-result write. A separate check proves
the actual rejected Save preserves configuration bytes. The corrected run is
**160 PASS / 56 FAIL**, with no harness exception and normal Excel closure, retaining
all 153 previous passing identities. The intentionally tightened sign-out assertion
is the only formerly passing new check that changes to RED. No runtime regression
or implementation GREEN is inferred from absent controls.

All 260 package pins, 15 protected sources and the reviewed Shipping visibility-only
change remain preserved; four changed scripts parse. Runtime/package/static/chain
baselines remain unchanged, and earlier native failures remain unresolved. Next
protect schema-1/schema-2/mixed-schema compatibility, stable StepId edits and stale
editor sequence/context binding before implementing Core expectation validation
and the shared Operations editor. Remaining evaluator integrity/storage, guides,
comparison and full Release 1 acceptance remain mandatory.

**Compatibility/editor-binding entry:** The D18 gate now adds actual library
selection of schema-1, schema-2 None and schema-2 expected-step records, plus mixed
journals, premature expectations, duplicate/unknown fields, invalid types/codes,
missing terminals and 257 steps. Variants preserve the real owning observations;
original disposable journal bytes are restored after every case. Real editor
actions test stable hidden StepIds on reorder, a stale editor across Stop/Start,
and nonempty draft clearing on sign-out. No new architectural behavior is proposed.
The first implementation boundary is strict headless Core expectation validation
and compatible recording reads, then the Operations-owned shared editor and
captured staging. The remaining evaluator/guide/release scope is unchanged.

**Expectation reader compatibility checkpoint (code `ad3f92f`):** The frozen candidate reports
173 PASS / 61 FAIL; the strict Core reader candidate
`deploy/validation-expectation-schema` reports 175 PASS / 59 FAIL, retaining all
173 passing identities and passing all 15 compatibility checks without a harness
exception. Two Core modules change and one validator is added. All five packages
build/compile, including Operations cold start; isolated full chain 31/31, live
roles 48/48 and Create Warehouse 15/15 pass. Tracked generated reports are restored
exactly and Excel is closed. All 265 package pins and protected sources remain
preserved. Static metrics retain all 28 large-module ratchets, 192 duplicate groups
and 45/9 unresolved/literal calls. One maintenance candidate is removed by existing
test reachability evidence; no code is deleted.

Core reads schema 1 and schema 2; writing schema 2 still requires the already-RED
shared editor and context/SequenceId-bound staging. Those are the next implementation
steps, not an optional alternative to the approved writer. The 59 remaining RED
checks, evaluator/storage, both presentations, guides and full Release 1 gates
remain open. No new editor layout or human acceptance is claimed, and prior native
failures remain unresolved. Controls v1.136 and the maintained recording results
record the same scope; Architecture v4.11 D18 is unchanged.

**Diagnostic result implementation entry:** D18's Evaluation wire refinement names
the already required journal, expectation, actor/policy, publication, matching and
source-group provenance fields without changing their meaning. Extend the actual
Evaluate-handler tests to require those bindings, final content hash, all four
terminal source references and exact applied line/key evidence before implementing
result writes. Core must retain the explicitly loaded publication and read the
capture policy's saved version; current visibility remains independently enforced.
Operations owns Evaluate and the read-only result. Missing product results are
expected RED, while fixture, compile and COM failures remain harness failures.
Preserve all 210 prior GREEN checks and the current package pins. Guide authoring,
both presentations, comprehensive coverage and full Release 1 acceptance remain
required; this is not a substitute for those outcomes.

The frozen expanded run reaches 151 PASS / 30 FAIL before the next processor call
loses Excel/RPC. Pending/Partial result-evidence RED is observed; Applied is not
reached. The terminal Harness.Exception is not product RED. The verified empty
residual closes normally, with file hashes preserved. The first evaluator candidate
builds/compiles all five packages and passes Operations cold-start dependency
validation; behavioral and regression acceptance are in progress. Additional
actual-library Refresh checks recompute hashes around deliberately malformed
disposable results to distinguish integrity from schema/journal-reference validity.
No native-failure fix or complete diagnostic acceptance is claimed.

**Corrected diagnostic candidate (2026-09-14, code `01e32ff` pushed; in validation):**
`deploy/validation-diagnostic-evaluation-reasons` completes the RED-only isolated
evaluation route at 185 PASS / 4 FAIL, retaining all 172 preceding passing identities.
The seven newly protected product defects are corrected; six timestamp assertions
are corrected to the existing millisecond UTC wire. The four failures are foreground
captures. Five packages build/compile with Operations cold start; compiled comparison
finds eight new Core modules and eight changed components, with no deletions. Static
evidence preserves all 28 existing module limits, 192 duplicates and 9/45 dynamic-call
counts. Verification preserves 310 package pins and protected sources; 24 scripts
parse and 83 local links resolve. These implement D18's approved evaluator contract.

The candidate passes full chain 32/32, live roles 48/48 and Create Warehouse 15/15,
with successful controller exit and exact settings/tracked-report restoration.
An Excel access violation in combase.dll during the chain still requires diagnosis;
passing chain assertions do not establish native reliability. Its verified empty
recovery instance closes normally after retaining recovery files for later viewing.
Three loaded add-in hashes remain unchanged. The subsequent complete packaged route
finishes at **293 PASS / 4 FAIL**, retaining all **210** previous passing identities
and gaining **83**, with no duplicates or terminal harness exceptions. All original
Viewer/recording/reader/editor/evaluator actions remain in that route. Library and
editor layouts, exact identities/unknown columns and unrelated workbook/file checks
pass; the four foreground captures remain failed. Excel closes normally. This run
does not resolve earlier native failures. The separate limits/storage gate passes
77/77, retaining every prior identity: 256 actions/512 observations, truthful partial
closure, continued ordinary action 257, exact 1 MiB save/read, identical append and
explicit oversize rejection with no partial publication. Its controller exits and
Excel closes normally. Final verification again preserves all 310 package pins and
protected sources, 28 module limits, 24 parsed scripts and 83 local links.
Result-selection fidelity, guide lifecycle, both presentations and remaining
comprehensive/visible/NAS/human Release 1 gates remain required. Slice 4be stays open.

**Next evaluator test-first work:** Exercise the actual library selection handler
during Evaluate, and Refresh after the same journal advances to a new version.
Require the displayed/saved result to retain its exact selected run/version binding;
the existing between-run draft checks do not cover reentrant evaluation. Also protect
complete terminal-reference retention on saved reads and D18's display timestamp
format through the actual result pane. Current `modEvaluationPresentation.Render`
prints wire timestamps directly; add the display RED before changing that renderer.
These tests constrain the approved D18 binding/provenance/display rules. They do
not authorize a new contract or substitute for guides, both presentations or UAT.

**Selection-continuity test entry:** D18 now makes the existing selection constraint
explicit at the append and UI-return boundaries. A changed context, selected journal
binding or staged expectation invalidates a pending evaluation; a result already
committed remains immutable but cannot populate a different selection. Same-version
Refresh retains the selected result, while a newer journal version clears it. Add
one-shot test callbacks at those two boundaries in disposable package projects,
dispatching the real list selection handler. Assert no stale append before commit,
no stale result attachment after commit, correct version refresh, full saved terminal
references and verified-UTC display before implementation. Preserve all 293 existing
passing checks and the 77/77 boundary gate. This is semantic inheritance of approved
D18, with no new capability or workflow execution contract.

The initial expanded full run terminates at 68 PASS / one Excel/RPC harness failure
before the new cases. The recording harness now installs probes once before fixture
and form activity, preserving all operator actions and check identities; counters
measure Viewer calls relative to completed fixture setup. The next run passes its
no-loaded-forms check, proves Pending evaluation/load timestamp display RED, and
then loses Excel/RPC during Partial (154 PASS / 3 FAIL, including one harness failure).
This does not establish native reliability. The frozen diagnostic route subsequently
completes at 197 PASS / 19 FAIL with no terminal harness exception and Excel closed.
Its fifteen behavioral failures establish RED for six timestamp displays, omitted
terminal references and eight stale append/attachment cases. Four foreground captures
still fail. The Core/Operations correction follows that observed RED: local selection
and intent revisions protect append, selected journal binding protects saved reads,
and the form rejects stale return values and clears results on a new journal version.
Same-version Refresh and immutable committed results remain protected. Build/compile,
full 293-check retention and Release1 gates remain required before completion.

**Diagnostic-pane visible-evidence entry:** Continue 4be.4 against the unchanged
`7c1bb5e` candidate and retain its 324/324 packaged baseline. Add
`-CheckEvaluationVisualEvidence` to the complete actual-handler route: capture saved
Pending, Partial and Applied results at minimum/default/larger/restored sizes and
with terminal sources scrolled into view. Verify selection/result/text and original
training/activity/publication bytes, with no publisher or Shipping-owner read calls.
This is evidence for D18's existing layout/visible/non-mutation requirements. No new
runtime contract or manufactured behavioral RED is needed for this test-only work;
an observed product defect must receive its own focused RED before correction.

**Library-title test entry:** D18's saved-run surface now explicitly names the
window title **Action Paths**, matching the existing public library name and
declared form caption. This is a semantic-inheritance clarification, not a new
workflow or authority contract. After the actual packaged library launcher,
assert the runtime Caption before any title correction. Capture discovery may
use that observed title while the separate approved-title assertion remains
binding. Keep every existing visual, selection and preservation check. No runtime
correction is authorized by a failed capture lookup alone.

The title checkpoint's Applied capture exposed a lifecycle notice defect: **Conclusion
observed** appears with **conclusion not evaluated**. Add its own actual-handler
RED before correcting the notice under D18's separate lifecycle/result rules.
Do not include that wording in a title-only completion claim.

The title's actual packaged RED completes 254 PASS / 6 FAIL: three generated
`UserForm1` titles and three foreground captures. The one-line Operations
initialization correction follows that RED. Its isolated `validation-library-title`
five-package build and explicit compile pass; only the compiled library form differs
among 225 compared components. The complete visual route finishes 363 PASS / 5 FAIL,
retaining all 324 baseline identities and passing all three title assertions.
The five failures are foreground captures, including two that passed during RED.
Excel closes. A test-only bounded activation retry retains strict owner/foreground
checks and is now under full-route validation on the unchanged candidate. No visible
gate, native repair, deployment or human acceptance is claimed from this attempt.

The retry completes **368/368 PASS** on the same candidate, retaining all 324
baseline, 254 RED-pass and 363 first-attempt passing identities. All fifteen pane
and four editor captures pass and are inspected. The title's RED/GREEN is complete;
the separate lifecycle-notice defect remained open at that checkpoint. Excel exits normally, with no
Event1000 recorded for that process; this does not establish a general native repair.
All 320 package pins/protected sources remain intact, with three scripts parsing,
83 links resolving and all 28 module limits holding. The same candidate's Release1
chain completes 32/32, live roles 48/48 and Create Warehouse 15/15. Tracked reports
and local settings are restored, all jobs terminate and Excel closes. Event1000
records `c0000409`, module unknown, during the chain; passing assertions do not
establish clean native execution or a repair. Guide authoring,
How-To/Diagnostic/Compare selection, export/import, deployment and human acceptance
remain required and are not inferred from this gate. The subsequent capture-only
notice checkpoint supplies its own actual-handler RED/GREEN.

**Capture-only lifecycle notice test entry:** Under D18's existing separation,
the library notice is **Stopped. Capture frozen.** both before and after evaluation.
The separate saved-result surface owns evaluation status. Extend the real packaged
selection/Evaluate checks for Pending, Partial and Applied before changing Core's
notice. Expect the current **conclusion not evaluated** suffix to fail; immutable
journals/results, policy, selection and every preceding 368 GREEN remain protected.
This semantic-inheritance clarification changes no recording/result schema or
authority and needs no new architecture decision. Controls v1.146 is synchronized.

The focused frozen-package RED completes 260 PASS / six expected notice failures.
The one-line Core correction then builds/compiles all five packages; only
`modRecordingReader` differs among 225 compiled components. Full GREEN completes
374/374, retaining all prior 368 and RED 260 passing identities. All nineteen captures
pass and are reviewed, including byte-identity verification for six restored/source
views. Static metrics and every component line count remain unchanged; all 325 package
pins/protected sources and 28 module limits hold. Excel closes with no Event1000 for
that test process; this is not a general native repair. The Release1 chain completes
32/32, live roles 48/48 and Create Warehouse 15/15. Local settings and all three
tracked reports are restored; all jobs terminate and Excel closes. No Excel Event1000
is recorded in the verified chain window. Earlier native failures remain unresolved.
The capture-only notice correction is complete for this isolated checkpoint, with
controls v1.147 recording the same evidence; full Slice4be/Release1 acceptance is open.

Runtime/tooling checkpoint: code `5e2c45a`, isolated candidate
`deploy/validation-recording-notice`. Continue with actual packaged recording
proof across the remaining roles; preserve the full D18 guide/comparison/import/export
and visible user-acceptance requirements.

**4be.5 guide draft entry D13 step:** Continue the approved authored-guide and
How-To/Diagnostic/Compare work after code 9bfac38. The normative guide draft entry
refinement names Create guide and the Operations guide editor, with Core validation
and exact source binding. A real Admin Save sequence, frozen through the actual
recording controls and selected in Action Paths, protects creation, reused editor,
authored-versus-observed wording, immutable source evidence, cancellation and
maintenance/context guards. Missing entry/editor behavior must be RED before any
runtime implementation. Test facades may invoke actual controls and inspect views;
they must not create the draft or simulate an owning result. Preserve the existing
packaged GREEN baseline and all 4be.5 save/version/search/presentation/comparison/
export/import and final acceptance requirements. This first gate covers draft entry,
not completed guide authoring or Release 1 acceptance.

The first focused route establishes 41 PASS / 14 expected missing-entry/editor
FAIL, after two fixture errors were corrected (required user argument and explicit
ACTION_PATH_MAINT grant in the disposable fixture). Both earlier runs preserve
76 existing Viewer/recording checks; their harness failures are not product RED.
The expanded focused route additionally protects actual authored step editing,
Move up/Move down/Remove with stable StepIds, unchanged original observations,
900 by 650 default / 760 by 600 minimum layout, and stale-evidence clearing.
`-GuideDraftOnly` retains packaged Viewer/source-fixture prerequisites while avoiding
the broader recording suite during focused iteration. `-CheckGuideDraft` retains
that suite for the combined regression gate. Runtime implementation remains pending.

Expanded draft RED is now verified: 41 PASS / 26 expected guide failures, retaining
all 55 preceding focused identities and 38 baseline Viewer GREENs. All five
instrumented compiles pass before forms. Runtime source, the five frozen package
hashes and both unrelated user documents remain unchanged. Excel closes; all four
attempts' audited windows have no Application events 1000/1001/1002. The code
`plan022_slice4be_guide_draft_results.md` records exact reports and fixture failures.
Next implement the captured Core draft boundary and Operations editor through
these protected handlers; this RED checkpoint does not accept any guide runtime.

Guide test checkpoint: code **7972cd0**, with controls v1.161. The frozen runtime
remains code 9bfac38 in `deploy/validation-event-detail-labels`.

Guide implementation is now in progress in an isolated candidate. Four new Core /
Operations components and the changed Action Paths form compile in all five
packages; the first behavioral run retains 42 PASS but fails 25 guide checks
because permitted entry remains disabled. This is not guide acceptance. The code
guide evidence records the two preceding compile defects and ongoing guard trace.
The explicit maintenance exception permits only the three reviewed duplicate
groups `041b258526c03dfd`, `b0153d52048d1613` and `e9b4f06801108607`: distinct
command-literal handlers and small typed editor ownership/cleanup procedures.
Feature growth is four components / 40 procedures / 491 lines; all 28 existing
module limits and 9/45 dynamic-call counts hold. The exception preserves typed
calls under D18 semantic inheritance and does not change normative behavior.
Required packaged lifecycle/expectation regressions remain pending.
The trace confirms policy serialization error 5 from Excel's Double catalog-version
projection. The new guide boundary normalizes the already-validated integer;
the shared serializer contract is unchanged. The corrected isolated candidate
compiles in all five packages and passes 72/72 focused checks: all 67 protecting
identities plus five actual editor captures, each directly reviewed. The final
static baseline is 236 components / 5,921 procedures / 130,591 lines, with the
reviewed three duplicate exceptions, 9/45 dynamic calls and all 28 existing module
limits retained. Ten current/frozen package hashes and both unrelated documents
are preserved. Eleven audited build/compile/test windows contain no Application
events 1000/1001/1002; build cleanup intervals precede no-Excel compile preflights.
Combined recording, detail, Viewer/Shipping state, Boxing/Shipping, evaluation and
full chain/live-role regressions are running serially. This is focused draft
evidence, not acceptance of persistence, presentation/comparison, transfer,
deployed/NAS behavior or human Release 1 UAT.
The combined recording/guide gate now passes 105/105 and Event Detail passes
34/34; both terminate at exit 0 with Excel closed. Exact reports are retained in
the code guide evidence. Remaining serial gates are pending at this checkpoint.
Focused GREEN implementation checkpoint: code **b10b301**, pushed to main;
controls v1.162. This is a reviewable implementation checkpoint with the remaining
gates and full guide contract still open, not a completed-slice declaration.

Subsequent unchanged-draft gates pass Viewer/filter/Shipping state 94/94,
evaluation 376/376, and Boxing/Shipping 1,707 PASS / seven known unapproved D8-A
FAIL, preserving all 1,714 Boxing/Shipping identities and 18 owner-return facts.
The full chain terminates at 5 PASS / one RPC harness FAIL; live roles are
32 PASS / one harness FAIL at canonical projection rebuild, Create Warehouse
15/15. All three tracked reports are restored. A verified empty residual Excel
session closes normally with visible recovery-file retention; terminal Excel
closure is verified. The complete chain window contains no Application events
1000/1001/1002, but HRESULT 0x800706BE remains unresolved. This candidate has no
passing full-chain gate. Exact reports and scope are in the code guide evidence.

**Next 4be.5 D13 step -- immutable Save guide:** The normative guide-save
refinement names `btnSaveGuide`, publication wording and the fixed `Guides`
storage child/schema. This is a discovered implementation detail under D18
semantic inheritance, not a new architecture exception. Before runtime changes,
exercise the actual packaged Save handler for first publication, a second immutable
version, retained StepIds/original observations, blank-name/oversize rejection,
source/context/policy/capability guards and unchanged recording/activity/authority
bytes. Missing Save control/version behavior must be meaningful RED against the
frozen guide-draft candidate. Protect the existing draft, recording, Viewer and
role GREENs. Headless Core owns validation and atomic guide publication; Operations
owns the captured editor. Initial fields and prior-version links must match the
normative schema rather than treating a guide as a recording journal. Save does
not evaluate or create business Events. Persistence remains unimplemented at
this test-entry definition; discovery/editing, guide expectations, all presentation
and comparison modes, transfer and real-user provisioning remain required.

The corrected Save guide test completes 97 checks: 71 PASS / 26 expected
missing-save and unclosed-source-label FAIL. All 67 prior Viewer/draft checks
pass, with no harness exception. Earlier attempts stop at a foreground-capture
guard and a stale fixture Viewer after re-sign-in; the fixture now reopens Viewer
under the new session and verifies restored capability/Auth bytes. Runtime remains
unchanged. Ten current/frozen package pins and both unrelated documents hold;
Excel closes, and all three test windows have no Application events 1000/1001/1002.
The code Save guide evidence records exact reports. This establishes the D13 RED
entry for implementation, not acceptance of persistence or the full guide contract.
Test checkpoint: code **e96863d**, pushed to main; runtime remains **b10b301**.

The isolated Save candidate now builds and explicitly compiles all five packages,
including Operations cold start. It adds two Core guide modules and changes only
the five declared draft/source, journal/evaluation-store and form components.
The fixed child-root check is shared to preserve the same Guides/Evaluations
target and reparse rules without duplicated storage-root code. Existing evaluation
GREEN is a required regression. Fresh static evidence adds two components/eight
procedures/299 lines; 194 duplicate groups, 9/45 dynamic calls and all 28 existing
module limits hold. The focused GREEN attempt is running; acceptance remains open.

The first behavioral Save run is 80 PASS / 17 dependent publication FAIL, with
all 67 preceding checks retained. RED-only instrumentation pinpoints the guide
validator's cross-namespace duplicate check. Existing `modActivity.BeginAction`
uses the generated ActivityId as its REQUESTED RecordId; those original values
must remain unchanged. The normative save refinement now explicitly preserves
that relationship while rejecting duplicate records/steps and source collisions
with new guide/Step identities. This clarifies the inherited activity schema;
it does not introduce a new identity scheme or weaken record uniqueness. Correct
the guide validator under the existing original-observation RED, then remove
diagnostic instrumentation before GREEN. Foreground capture attempts also fail
independently and remain unproven; do not claim visible acceptance from compilation.

The corrected `deploy/validation-guide-save-identity` candidate now passes **97/97
focused GREEN**, retaining every protecting RED identity. First/next versions,
original observations, source/package provenance, stable authored steps, rejection
paths and binding guards pass through the real controls. Four layouts, all five
explicit compiles and Operations cold start pass. Static growth is two components,
eight procedures and 301 lines; duplicate groups remain 194, dynamic calls 9/45,
and all 28 module limits hold. Twenty current/frozen package hashes and both
unrelated documents are preserved. The terminal run closes Excel, and ten audited
build/compile/test windows contain no Application events 1000/1001/1002. Exact
evidence is in the code Save guide record. Current captures, the shared-root
evaluation regression, all other required regressions and full guide/Release 1
acceptance remain open. This is a focused implementation checkpoint, not a
completed Slice 4be declaration or deployment authorization.
Focused implementation checkpoint: code **c3bb22b**, pushed to main. The full
evaluation/visible-library regression passes 376/376 against the unchanged
isolated candidate, retaining all prior identities and protecting the shared
folder helper. All 15 diagnostic captures were directly reviewed across pending,
partially applied and applied states and supported layouts/source viewport. The
controller exits 0, Excel closes, five candidate hashes and both unrelated user
documents are preserved, and the run window has no Application events
1000/1001/1002. The code Save evidence records the exact report and scope. Save
editor captures and all remaining release gates stay open.

**Next 4be.5 D13 entry -- published-guide discovery/read:** Apply D18's
published-guide reader refinement before runtime edits. Operations adds
`btnPublishedGuides` to Action Paths and the captured reusable
`frmActionPathLibrary`; headless Core owns exact-version/hash/predecessor and
current-policy validation. The packaged test must create both guide versions
through the existing Save handler, then exercise real entry, search, selection,
Refresh and Close controls. Protect ordinary-reader access, original captions,
authored versus observed ordering, immutable source bytes, current-policy and
target invalidation, corrupted/missing predecessor rejection and all four layouts.
Missing reader behavior is RED; a broken Save fixture or absent test seam is not.
Retain all 97 Save/draft/Viewer checks. No runtime reader implementation or GREEN
is claimed by this entry. Both presentations/comparison, direct event curation,
editing/expectations, transfer, provisioning and all remaining Release 1 gates
remain required; D8-A remains unapproved.

The published-guide test now establishes **102 PASS / 30 expected missing-reader
FAIL**, 132 total, preserving all 97 prior Save/draft/Viewer GREEN identities.
The unchanged candidate compiles with its probes and completes without a harness
exception. Reader source/read-only preservation checks pass. Terminal exit is 1
as expected, Excel closes, 20 package hashes and both unrelated user documents
remain unchanged, and the run has no Application events 1000/1001/1002. The code
published-guide evidence records the exact report. Implement the declared reader
through typed Operations handlers and headless Core reads next, retaining this
RED set. Runtime, static baseline and accepted deployments remain unchanged at
this test-first checkpoint; it is not reader GREEN or full Slice 4be acceptance.
Test/evaluation checkpoint: code **627a961**, pushed to main; runtime remains
**c3bb22b**. No accepted deployment or user workbook was changed.

Reader implementation now exists in isolated `deploy/validation-guide-library`.
Five-package build/explicit compile and Operations cold start pass. Two existing
components change (`modActionPathRead`, `frmActionPaths`); three are added
(`modGuideLibraryRead`, private `modTrainingReadContext`, `frmActionPathLibrary`).
The private helper preserves the existing recording guard while sharing it with
guide reads. Static growth is three components / 22 procedures / 345 lines;
duplicate groups decrease 194 to 193, dynamic calls stay 9/45 and all 28 existing
module limits hold. The focused reader gate passes 132/132, retaining every RED
identity. Terminal exit is 0 and Excel closes; all 25 current/frozen package hashes
and both unrelated documents remain unchanged. Build/compile/test windows contain
no Application events 1000/1001/1002. The code reader evidence records the exact
report and cleanup scope. The dedicated actual Save/reader capture run passes
135/135, retaining all 132 preceding identities and three directly reviewed
captures. Publication wording/version labels and authored versus observed panes
fit visibly. Excel closes, 25 package hashes and both unrelated documents are
preserved, and the capture window has no Application events 1000/1001/1002.
Refreshed static evidence retains the same metrics and all existing limits.
Serial recording, detail, Viewer/Shipping state, Boxing/Shipping, evaluation and
full-chain regressions are running; verify each result before acceptance. No accepted
package or guide schema changes, and no full Slice 4be acceptance is claimed.
Focused reader implementation checkpoint: code **fdcfac9**, pushed to main;
the isolated candidate remains under serial regression with all remaining gates open.

Completed reader regression gates on 2026-09-15: combined recording/guide/Save/
reader **170/170**, retaining all 105 preceding combined identities; Event Detail
**34/34**; Viewer/filter/Shipping state **94/94**, retaining their preceding
identities. All three exit 0 with Excel closed, unchanged candidate/user-document
hashes and zero Application events 1000/1001/1002 in their measured windows.
The maintained published-guide evidence records exact reports and UTC windows.
Boxing/Shipping, evaluation and the full chain still require separate terminal
verification. No D8-A approval, deployment or full acceptance is implied.

Subsequent Boxing/Shipping terminal verification retains all 1,714 preceding
identities and 18 owner-return observations: **1,707 PASS / seven unchanged D8-A
FAIL**. Immediate cleanup is pending in its exit snapshot; the next evaluation
preflight verifies no Excel, and that extended window has zero Application
events 1000/1001/1002. Candidate and unrelated-document hashes are unchanged.
Evaluation/full-chain results and direct Boxing-capture review remain open.
All 22 Boxing-run PNG captures have subsequently been directly reviewed. They
show actual accepted/rejected Make/Unbox, tracking-disabled/unavailable feedback,
recorded observations, published details, Settings and permission denial. The
stale-session capture is only partially painted; do not treat it as complete
visible operator feedback. Evaluation/full-chain outcomes remain pending.
Six detail/Viewer captures have now been directly reviewed. The long Coverage
value is clipped at Event Detail's default/restored width and readable maximized.
Track default-width long-field readability as remaining visible acceptance work;
protect it with an actual form-action test before any presentation correction.
This discovery does not replace the active guide-expectation D13 entry below.

The completed evaluation gate is **376/376**, retaining all preceding identities.
All 15 diagnostic captures are directly reviewed; partial application awaits the
remaining source and only all-applied evidence yields Conclusion observed. Excel
closes and the measured window has zero Application events 1000/1001/1002.
Subsequent **full chain fails: 5 PASS / one harness FAIL**, with live roles
**14 PASS / one harness FAIL** at **Run Receiving ConfirmWrites**, `0x80020009`.
Create Warehouse is **15/15**. This is not acceptance or guide-expectation RED.
The chain window also has zero Application events 1000/1001/1002; the cause remains
unproven. Its remaining Excel process is verified empty before normal Quit; the
inspected recovery dialog retains files for later viewing. No forced termination
occurs. Local settings and the three tracked reports restore exactly, and all
25 frozen package hashes plus both unrelated documents remain unchanged. Exact
reports/windows and separate visible limitations are in the maintained reader
evidence. Continue independent guide-expectation RED on the unchanged candidate;
the full-chain failure still requires investigation before release acceptance.

**Next 4be.5 D13 entry -- guide expectation authoring:** Follow D18's new
guide-expectation refinement. The existing shared expectation editor gains an
explicit guide-draft scope, entered through `btnGuideExpectedConclusion` with
`lblGuideExpectationSummary`. Use for this guide stages intent; only Save guide
publishes the definition with its immutable version. Initial None, current
source/context/policy/ACTION_PATH_MAINT checks, Cancel, exact prior links and
original observations remain binding. Preserve the editor's existing dimensions.
Establish packaged RED for actual entry, scope wording, initial None, registered
ordered/repeated steps and retry choices, cancellation, staging without writes,
first/next saved definitions and invalidation before runtime changes. Preserve
all 132 reader/draft/Save/Viewer checks. Current reader regressions continue on
the frozen candidate independently. No expectation implementation or RED result
is claimed by this entry. Guide-bound evaluation, both presentations/Compare,
closed-guide editing, direct event curation, transfer and full acceptance remain.
Prepared test/evidence checkpoint: code **7c11619**, pushed to main; runtime
remains **fdcfac9**. Four PowerShell parsers, 89 local links, 25 frozen package
hashes and both unrelated-document hashes pass. Guide-expectation RED is queued
until the existing serial regression runner releases Excel.

**Guide-expectation RED verified:** The unchanged reader candidate completes
**136 PASS / 30 expected FAIL**, 166 total, retaining every preceding 132
reader/Save/draft/Viewer identity. Existing captured-expectation and immutable Save
fixtures succeed; there is no harness exception or duplicate identity. Missing
actual guide entry/scope, summary, stage/cancel/invalidation and persisted intent
are the behavioral failures. Four new preservation cases pass. Excel closes;
all 25 frozen package hashes and both unrelated documents are unchanged, and
the UTC 2026-09-15 08:52:30.7577605--09:00:26.0713116 window has zero Application
events 1000/1001/1002. Exact report:
`reports/runtime/slice4be-viewer-published-read/1515cc3719c745afa21fa5452e43c594/red.json`
in the code repository. Implement the approved guide-draft scope and immutable
expectation saving next; full-chain failure and all broader release work remain.
Verified RED/evidence checkpoint: code **a5babcb**, pushed to main. Runtime
remains **fdcfac9**; four PowerShell parsers, 89 local links, 25 package hashes
and both unrelated documents are verified. Excel and both gate runners are closed.

**Guide-expectation implementation GREEN:** The isolated
`deploy/validation-guide-expectation` candidate passes **166/166**, preserving
every protecting RED identity and all preceding 132 reader/Save/draft/Viewer
checks. Core holds independent guide intent; the shared Operations editor names
its explicit scope. Use stages only, Save publishes immutable definitions, and
Cancel, recording lifecycle, scope switching, policy/capability/target loss and
source/identity preservation pass. All five packages build/compile; Operations
cold start passes. Compiled comparison retains 234 components with exactly seven
changed and none added/removed. The behavior controller exits 0 with immediate
Excel closure false; normal closure is verified separately at UTC 2026-09-15
09:24:45.4286630. Report:
`reports/runtime/slice4be-viewer-published-read/ff91f4fc88104eeda788a6cbe13c2811/green.json`.
An earlier combined capture gate fails after 81 passing checks at the first
Save-guide image's foreground guard, producing no PNG. This is a harness failure,
not product RED. A test-only diagnostic preserves the guard and records window
identities/classes without captions or workbook values. The separate diagnostic
capture gate also ends **81 PASS / one harness FAIL** with no PNG. All three
attempts report activation success but retain VS Code's foreground root instead
of the intended Excel form. This identifies the competing window, not a proven
cause or correction. Excel closes and its measured window has zero Application
events 1000/1001/1002. Report:
`reports/runtime/slice4be-viewer-published-read/fa883641b2e6485faaa55661ab440d53/green.json`.
Refreshed static evidence is 241 components / 5,958 procedures / 131,327 lines;
duplicate groups decrease 193 to 192, dynamic calls remain 9/45 and all 28
existing module limits pass. All 30 candidate/frozen package hashes and both
unrelated documents are preserved; four scripts parse and 89 local links resolve.
Build/compile/behavior native windows have zero Application events 1000/1001/1002,
including observed cleanup. Full visual acceptance, broader candidate regressions
and full-chain success remain open. The maintained guide-expectation evidence
records commands, exact reports and cleanup observations. The next independent
gate is the combined recording/guide/expectation regression on this candidate.
Focused implementation checkpoint: code **827a427**, pushed to main. No accepted
deployment, installed package or NAS runtime changed; the slice remains open.

**Combined guide-expectation regression GREEN:** The same frozen candidate
passes **204/204**, retaining all 170 preceding combined identities and all 166
focused protecting identities with no duplicates. Report:
`reports/runtime/slice4be-viewer-published-read/a43f12170584430086b493cd222958a5/green.json`.
Exit 0; immediate Excel closure is false, with normal closure verified at UTC
2026-09-15 09:45:20.7984561. The extended 09:32:51.6975893--09:45:20.7984561
window has zero Application events 1000/1001/1002. All 30 package hashes, both
unrelated documents and the existing compile/static scope remain verified.
The blank-form capture comparison succeeds with both hidden and visible Excel;
hidden Excel alone therefore does not reproduce the packaged foreground failure.
All three images are reviewed, exposing an incomplete restored-hidden rendering
despite capture success and full dimensions. The next packaged diagnostic keeps
Excel visible from startup under an explicit test-only flag, preserving all
capture guards. It stops **9 PASS / one harness FAIL** at Viewer startup, before
guide behavior: an inspected Visual Basic dialog reports Automation error 440.
Debug identifies `modInventoryViewer.OpenInventoryViewer` at the ordinary
`mInventoryViewer.SetWarehouse warehouseId` call, not an underlying cause.
Reviewed VBE Reset ends the disposable execution; the caller then returns
`0x800A9C68`, and normal cleanup closes Excel without forced termination.
Its UTC 2026-09-15 09:47:14.7141703--09:57:27.9017095 window has zero Application
events 1000/1001/1002. Report:
`reports/runtime/slice4be-viewer-published-read/5cada388ce034b26b4ca0a771ef9990e/green.json`.
The stopped Viewer call must be investigated before another capture attempt or
runtime correction. No foreground fix or visible acceptance is established.
The full-chain harness also adds reporting-only stage labels to distinguish
Receiving form invocation, outcome/status reads, capability revalidation and
projection inspection. Workflow calls/assertions are unchanged; a later chain
run must supply any new location or result. The maintained capture-diagnosis
record contains exact commands, observations and the remaining limitations.
Six relevant PowerShell scripts parse, 91 local links resolve, and all 30 frozen
package hashes plus both unrelated documents remain unchanged. Runtime remains
**827a427**; these diagnostic additions change no architectural contract.
Combined regression/diagnostic checkpoint: code **5b0505e**, pushed to main.
Excel and both controllers are closed; isolate the observed visible Viewer
startup call next, preserving the focused/combined GREEN and all release scope.

**Viewer startup measured comparison:** The harness assigns Excel.Visible=True,
but the default-startup focused trace reads Boolean False and opens/reuses the
visible modeless Viewer (15/15). Early-visible setup reads Boolean True and fails
at SetWarehouse's caption assignment (14 PASS / two startup FAIL). A saved,
reopened ordinary workbook preserves its active identity, Saved state and closed
file hash, but also fails there (16 PASS / two startup FAIL). The earlier 166/166
and 204/204 did not measure application visibility at callback entry; retain their
default-startup scope without a retrospective hidden/visible application claim.
All three compile five disposable instrumented packages, use actual public entry
and real Admin Settings/publication fixtures, and preserve captured context,
warehouse bytes, publication/Shipping authority and normal cleanup. The fixed
trace identifies the caption boundary, not its underlying cause. Extended
Initialize/Activate/Layout tracing and the reduced Viewer-only probe set each
report 16 PASS / two startup FAIL: initialization returns, with no Layout or
Activate entry before the caption failure. Recording/guide probes are not needed
to reproduce it. All nine startup attempts close Excel and have zero Application
events 1000/1001/1002 in their measured windows. Eight scripts parse, 91 links
resolve, and all 30 package hashes plus both unrelated documents are preserved.
Compile, insertion, visibility assumption and locked-file hashing setup failures
remain non-product failures. Exact reports and limitations are in the maintained
capture-diagnostic evidence. The saved-probe comparisons below supersede the
then-planned blank form caption comparison in the same session. No runtime,
normative contract or accepted deployment change is made.
Startup diagnostic checkpoint: code **cda7ecb**, pushed to main; runtime remains
**827a427**. Four invalid diagnostic option combinations are rejected before
Excel setup. Excel and both new diagnostic controllers are closed.

**Saved probe package comparison (2026-09-21):** Blank-form calibration isolates
the failed caption access to configuration workbook closure; showing the form
first or omitting workbook hiding does not prevent it. Calibration can change
whether the later Viewer opens, so it remains explicitly opt-in and is not a
startup acceptance substitute. With no calibration, writable temporary copies
left unsaved reproduce error 440 (18 PASS / two startup FAIL). Compiling/saving
the same probe set in temporary copies passes 20/20 through actual public Viewer
entry/reuse, with visible Excel, captured context, package bindings, warehouse
bytes/counters and saved-workbook preservation. Interpret the earlier failure
with its unsaved instrumentation state; no runtime caption/configuration patch is
justified by these results. SavedCopies is available for the compiled guide-
expectation gate. After metadata-reader harness corrections, 83 checks pass before
the first foreground capture fails. Retaining a saved/reopened ordinary workbook
gives 84 PASS / two FAIL: foreground still fails despite visible Excel and an
enabled/non-minimized form, and its in-memory Saved-state check fails. The exact
workbook path and closed-file hash remain unchanged; that state change's cause is
unproven. Neither attempt produces a PNG, and both controllers close Excel.
Full-chain regression is next; capture/workbook-state diagnosis remains open.
Frozen inputs and accepted deployment remain unchanged. An interrupted shown-form attempt has no final
result or proven normal cleanup; four temporary fixture references remained in
local settings and its original in-memory snapshot was unavailable. Preserve
that restoration limitation rather than guessing the prior selection. The
maintained capture evidence records exact reports and commands. No normative
contract changes; remaining feature, regression, physical and human gates stay open.
Saved-probe/capture checkpoint: code **66f909d**, pushed to main; runtime remains
**827a427**. The later result follows.

**Expectation candidate visible/chain evidence (2026-09-21):** The visible guide
gate passes 177/177, retaining all 166 focused identities and nine reviewed guide,
reader and expectation captures. Every capture satisfies the native foreground
guard with zero accessibility fallback attempts. An experimental fallback was
discarded after its label-only blank calibration target could not receive focus;
it is not credited as the cause of the passing captures. No runtime, normative
contract, global focus policy or accepted deployment changed.

The first chain stops 5 PASS / one harness FAIL, with live roles 32 PASS / one RPC
failure at projection rebuild and Create Warehouse 15/15. After verified empty-
instance normal cleanup and reviewed recovery-file retention, a clean comparison
passes 32/32, retaining all preceding identities, with live roles 48/48 and Create
Warehouse 15/15. The successful chain also requires assisted empty-instance normal
cleanup with recovery files retained. Both controllers restore all three tracked
reports and close Excel; their measured windows contain no Application events
1000/1001/1002. This is a functional chain pass, not proof of native repair or
unattended recovery. Static generation retains 241 components, 5,958 procedures,
1,184 scanner candidates and no new warning paths (27 before/after).

All 30 frozen packages and both unrelated documents remain unchanged. Event Detail
passes 34/34 with three reviewed captures in the experimental-helper comparison;
its default/restored Coverage clipping remains open. The retained native helper
passes Viewer/filter/Shipping state 94/94 with every prior identity. Its activity
and filter images are readable; the Shipping image is partially painted and does
not establish complete operator feedback. Boxing/Shipping preserves all 1,714
identities at 1,707 PASS / seven unchanged D8-A FAIL, with normal Excel closure
and zero matching Application events. All 22 images were reviewed; this run's
stale-session feedback is fully painted. Coverage clipping and adjacent Shipping
System Key/Alternative headings remain readability limitations. Evaluation passes
376/376 with every prior identity retained, normal Excel closure and zero matching
Application events. All 19 captures were reviewed: 15 diagnostic and four editor
views. Exact reports and limitations are in the maintained guide-expectation and
capture evidence. Full Slice 4be, guide-bound evaluation, both presentations and
Compare, editing/curation/transfer, physical deployment and human UAT remain open.
Visible/chain evidence checkpoint: code **ba7cfd0**, pushed to main; runtime
remains **827a427** and the retained test helper is **66f909d**.

**Event Detail overflow D13 entry (2026-09-21):** D18 now explicitly constrains
the existing locked field list to make complete permitted single-line text reachable through
its native horizontal scrolling at minimum/larger/maximized/restored sizes.
This restores required detail readability without a new user selection or
tracking contract. Extend the existing packaged Viewer-selection test first:
the fixed-width Coverage value must fail the full-text geometry assertion before
changing `frmEventDetail`. Preserve every prior detail check, cached profile and
source values, read-only behavior and unknown-column/source-byte checks. Require
focused RED/GREEN, isolated package/compile/layout/static gates and directly
reviewed scrolling evidence. The unchanged candidate establishes RED at 37 PASS /
four expected FAIL, preserving all 34 prior detail checks. Default/larger/restored
text capacity fails; native-maximized capacity passes. No harness failure,
duplicate, policy/data reload or native Application event occurs; Excel closes
normally. See [detail overflow evidence](../../invSys_fork/tests/integration/plan022_slice4be_detail_overflow_results.md).
The isolated correction adds one private helper/18 source lines and changes only
`frmEventDetail` among 234 compiled components. Build and all five package compiles
pass. Focused behavioral GREEN is 41/41, preserving the preceding 34 identities;
Viewer regression is 94/94. Duplicate-body groups and dynamic-call counts remain
192 and 9/45, with all 28 preceding oversized-module limits satisfied. Candidate
Boxing/Shipping retains all 1,714 identities at 1,707 PASS / seven unchanged D8-A
FAIL; all 22 captures are reviewed and the native Application event count is zero.
The immediate terminal Excel closure was False; the strict closure guard passed
before the next gate started 171 ms later, without intervention. Two later
41/41 capture runs show the scrollbar but do not prove movement; a third bounded-press
41/41 run has the same visual limitation. Evaluation regression retains all 376
identities at 376/376, with normal closure, zero matching Application events and
all 19 captures directly reviewed. Visible scrolling
and multiline rendering remain open. Candidate full-chain runs remain 5/1, with
live-role results 36/1 and 32/1 and Create Warehouse 15/15. Joining reordered harness
sections with explicit newlines restores the Boxing stage assignment without
changing any of the 59 macro calls; it corrects reporting, not the RPC failure.
The corrected-stage run fails during canonical projection rebuild. Recovery files
were retained, verified empty Excel processes closed normally, and all three tracked
reports restored. At that checkpoint, visible and full-chain acceptance remained open.

The subsequent macro-diagnostic run retains all **32 chain / 48 live-role / 15
Create Warehouse** identities and passes, with zero matching Application events.
It is an assisted functional pass: normal Quit of a verified empty test Excel and
directly reviewed retention of recovery files were needed for final cleanup.
The earlier empty process had already closed before Quit; no action was applied to
it. All three tracked reports are restored. The immediate post-input detail capture
still shows no horizontal movement; visible scrolling remains open. The chain used
frozen overflow XLAMs, while its static scan included the guide-binding source draft.

**Next 4be.5 D13 entry -- guide-to-run expectation binding:** Follow D18's explicit
binding refinement. First protect the actual Published guides -> **Use for selected
run** -> separate **Evaluate** handlers. Existing Admin recording and guide-authoring
handlers must create two immutable guide versions and a different observed run.
Expected RED is the absent apply control/guide binding, never a fixture or compile
failure. Preserve all prior guide/expectation checks; prove exact version/hash and
expected StepIds, selected-run ActivityIds, ordinary-reader access, no inferred
source-run success, stale selection rejection, context/policy/integrity invalidation,
explicit analysis scope and non-mutation before Evaluate. Both presentation methods
and Compare remain required after this bounded binding step. Event Detail's open
visible/full-chain gates remain separate and are not waived by this next test entry.
The opt-in `-CheckGuideEvaluation` gate is connected with 37 protecting checks;
its frozen-candidate RED is **171 PASS / 33 expected FAIL** across 204 unique
checks. All preceding 166 identities pass; every failure belongs to guide binding.
Excel closes normally after the immediate terminal observation, and the extended
window has zero matching Application events. See
[guide-binding test entry](../../invSys_fork/tests/integration/plan022_slice4be_guide_evaluation_results.md).
The isolated guide-binding candidate now builds and all five packages compile;
six existing compiled components change and one private Core helper is added.
Its static duplicate/dynamic-call counts and all 28 preceding module limits pass.
Packaged GREEN is **204/204**, preserving every RED identity, all 166 preceding
checks and all 37 new guide-binding checks. Excel closes normally after the immediate
terminal observation, with no intervention or matching Application event. The
visible packaged gate now passes **214/214**, preserving all focused 204 and prior
visible 177 identities. All 15 images were directly reviewed, including exact
guide/run provenance at minimum/default/larger/restored sizes and the separate
evaluation result. Normal Excel closure and zero matching Application events are
verified. Two earlier capture failures remain explicit; guarded native caption
input and temporary restoration of the owned form's topmost state support the
strict foreground check without changing runtime or global focus policy. Broader
candidate regressions, both presentations/Compare and human acceptance remain open.
The candidate Viewer gate now passes 94/94; Boxing/Shipping preserves all 1,714
identities at 1,707 PASS/seven unchanged D8-A FAIL. Both close normally with zero
matching Application events. All 25 captures were reviewed. Individual-file
inspection corrects an earlier batch-review filename association: Boxing
zero-quantity rejection and accepted-Unbox feedback are correct. Unpainted Shipping
regions remain a visible-evidence limitation; diagnose painting/capture timing
before proposing a runtime fix. Evaluation now passes **376/376**, preserving all
prior identities; all 19 images were individually reviewed, with immediate normal
closure and zero matching Application events. The candidate chain fails **5/1**,
live roles **32/1**, at processor automation during projection rebuild with RPC
`0x800706BE`; Create Warehouse passes **15/15**. A verified empty test Excel
instance needed normal Quit and reviewed recovery retention. Tracked reports are
restored; the zero-event audit does not identify the native cause. Keep the chain
gate open. Independent paired-view RED can run after verified cleanup.
The strict owned-foreground Viewer capture rerun passes 94/94 with normal closure
and zero matching Application events. All three images were individually reviewed;
Shipping Hold/Not Shipped feedback is fully painted without a runtime change.
Adjacent headings and the broader denied/stale Shipping capture gates remain open.
The separate visible-host Event Detail test passes 41/41 with seven reviewed
images, normal closure and zero matching Application events, but still no visible
horizontal movement. Compare the disposable list's locked/unlocked interaction
before proposing any runtime lock-contract change.
The same-input disposable comparison now isolates the lock property: the locked
list stays at the left edge; the temporarily unlocked list reaches the right edge.
Field values are unchanged and the original lock is restored; the run retains
41/41 checks and closes normally. Architecture D18 now records an explicit
**pending-approval** proposal to permit an unlocked but non-editable ListBox for
navigation. Do not implement that property change until approved. After approval,
add a native-scroll behavioral RED and prove read-only values/source preservation
in GREEN; the current geometry checks alone do not protect interaction.
The native-input RED is now **42 PASS / one expected movement FAIL**, with all 41
prior checks passing, unchanged displayed values, normal closure and zero matching
Application events. Its image detector was first validated against the reviewed
stationary/moving comparison. The proposal still awaits approval; do not infer
approval from the completed RED or the separate paired-view GREEN.

**Next 4be.5 D13 entry -- paired How-To/Diagnostic views:** D18 now names an
additional Operations **View guide and run** surface for the explicitly staged
exact guide and selected recording. Preserve the existing library's visible
observations and separate Evaluate. The new view uses the saved effective method
on fresh opening; local switching/Refresh retains the pair without saving a
preference or evaluating. Settings remains the preference writer. Headless Core
validates current policy, exact guide/journal and any displayed saved result.
Protect the actual entry/method/Refresh/Close handlers with a focused test on the
frozen guide-binding candidate before implementation. Expected RED is absent
controls/presentation behavior, not fixture/compile/focus failure. Retain all 204
guide-binding identities and test all three views, saved preference, no implicit
evaluation or source-run substitution, stale-pair rejection, current visibility,
readable layout and non-mutation. No paired-view runtime or acceptance is claimed
by this test entry; the guide-binding regression/chain gates remain separate. The
packaged RED now verifies **207 PASS / 36 expected FAIL**, with all 204 prior
identities retained, normal Excel closure, unchanged candidate hashes and zero
matching Application events. All failures are new presentation checks. The Core
read boundary and reusable Operations view now pass focused **243/243 GREEN**,
including all 39 presentation checks and 204 previous identities. Excel closes
normally after slow shutdown, with no intervention or matching Application event.
The candidate builds and all five packages compile; its compiled
scope is one changed library form plus the new Core reader and Operations view.
Static duplicates/dynamic calls and all 28 prior module-size limits hold. These
gates do not establish visible or full presentation acceptance. See
[paired-presentation test entry](../../invSys_fork/tests/integration/plan022_slice4be_guide_presentation_results.md).

The first expanded paired visible run retains all 243 focused identities with
272 PASS / one harness failure after the missing-step scenario. A generic control
call is rejected by Excel; cleanup receives `0x800AC472`. Excel closes normally,
with zero matching Application events and unchanged candidate hashes. The exact
control/action was not recorded. Preserve the failed run; add safe harness-stage
diagnostics for the retry without replaying a potentially mutating action. This
does not establish product RED or visible acceptance.

The retry stops earlier: 255 PASS / one harness failure, after the saved Diagnostic
preference check. It reaches all 204 guide-binding and 233/243 focused identities;
ten focused checks are not reached. All 25 captures were individually reviewed.
Excel closes normally, with zero matching Application events and unchanged
packages. Record the exact control/action and HRESULT chain at the shared harness
boundary next; capture only an already-foreground owned permitted form, without
focus changes or replay. The historical unavailable-step and dirty-Config cases
extend tests of existing D18 behavior through existing form handlers. No new
runtime contract or acceptance is implied. Independently, Viewer/filter/Shipping
state passes 94/94 with immediate closure, zero matching Application events and
all three captures reviewed. Boxing/Shipping, evaluation and full chain remain
separate requirements; the pending Event Detail scrolling proposal is unchanged.

The frozen paired candidate's Boxing/Shipping regression now retains all 1,707
preceding GREEN identities with exactly seven pending D8-A findings. Excel closes
normally after slow shutdown; zero matching Application events and unchanged
package hashes are verified. All 22 captures were individually reviewed, including
painted permission-denied/stale-session Shipping notices after the capture-only
owned-form helper correction. The offline failure-diagnostic privacy/no-replay
checks pass 10/10. Neither gate approves D8-A, the scrolling proposal, the failed
paired visible run or full Release 1 acceptance. Evaluation and full chain follow
serially before the next paired visible diagnostic.

The paired candidate's evaluation regression now passes 376/376 with every prior
GREEN identity retained. Excel closes immediately and normally; zero matching
Application events are verified. All 19 captures were individually reviewed and
hash-verified: pending/partial publication remains Awaiting, only the applied case
shows Conclusion observed, and the ordered expectation editor and its three sizes
remain readable. The full chain is next, followed serially by the paired visible
diagnostic. Neither pending gate is accepted by this regression result.

The paired candidate's full chain subsequently passes 32/32, ordered live roles
48/48 and Create Warehouse 15/15, retaining all preceding chain/live identities.
All three tracked generated reports are restored byte-for-byte; zero matching
Application events and unchanged frozen packages are verified. Cleanup is assisted:
native-owner inspection proves an integer workbook count of zero before normal
Quit, then the observed recovery prompt's retention option is selected, reviewed
and confirmed. No forced termination or recovery deletion occurs. This is not an
unattended shutdown claim or acceptance of the remaining paired visible/human gates.

The expanded visible diagnostic records 294 PASS / one harness failure, retaining
all 243 focused identities. Normal closure, zero matching Application events,
unchanged packages/test sources and all 39 individually reviewed capture hashes
are verified. Missing, extra, rejected, None, capture-off and current visibility
cases pass, including minimum-size bottom-viewport identities. The final historical
gap fixture incorrectly demands Stopped after an excluded command; the existing
recorder interruption path produces Incomplete / TRACKING_UNAVAILABLE. Preserve
the failed run, correct the fixture without runtime changes, validate saved flags,
and retain sanitized counts/Booleans on any further preparation failure. The retry,
unavailable-policy cases, paired preference restart and human acceptance remain open.

The corrected full visible retry now passes 305/305, preserving all 294 preceding
passes and all 243 focused identities. All 11 new availability checks pass, including
the actual Incomplete / TRACKING_UNAVAILABLE closure, both unavailable expected steps,
dirty-Config content clearing, unsaved/stored-byte preservation and restoration of
the exact historical result. Excel closes normally without intervention; zero matching
Application events and unchanged packages/test sources are verified. Visible evidence
combines 39 earlier individually reviewed images with three new availability images
from this same frozen candidate; all reviewed hashes are checked, not all retry images
claimed reviewed. Regenerated maintenance metrics and all 28 prior size limits hold.
The prior failures remain evidence; the COM rejection's root cause is not established.
Next protect fresh-process paired preference through the actual Settings and view
handlers. Closed-guide editing, direct curation, transfer, remaining D18 and human
acceptance are still required; neither pending architecture proposal is approved.

The first paired preference restart attempt preserves all 294 prior passes and
adds three passing pre-restart checks: Operations Save My Preference, Compare both
on open and an unsaved How-To switch. It ends at 297 PASS / one harness failure
because hashing the writable probe XLAMs encounters Excel's sharing lock before
restart. Normal cleanup, zero matching Application events and unchanged running
sources are verified; the one new How-To capture was directly reviewed. This is
not product RED or a fresh-session result. Move probe hashing after normal exit
and run the focused actual-handler restart fixture on the same frozen candidate;
retain the earlier regression evidence. No runtime change is needed for this
harness failure.

The focused retry now passes 27/27, including five instrumented package compiles,
four actual-handler fixture checks and all 15 restart checks. Normal owned-instance
closure is verified before a different native Excel process loads Operations with
Core/Domain and no Admin. Settings and the explicitly paired view restore the saved
Compare both preference rather than the unsaved How-To switch. Fresh entry infers
no pair or evaluation; Config/training/probe bytes remain unchanged. Final closure
is normal with zero matching Application events and unchanged running sources.
All three new captures were individually reviewed and hash-verified. The 305/305
visible gate and earlier preserved 294 identities remain separate regression
evidence; this small focused gate does not replace them. No runtime change was
needed. Continue the specified published-guide edit RED below; human acceptance,
direct curation, transfer and other D18/Release 1 scope remain open.

**Next 4be.5 D13 entry -- reopen a published guide for editing:** Follow the
normative published-guide editing refinement. Operations Published guides adds
`btnEditPublishedGuide`, **Edit guide**, with ACTION_PATH_MAINT checked in Core.
The exact selected ID/version/hash supplies authored fields, stable authored and
expected StepIds, observations and SourceRun; a selected diagnostic run supplies
none of those values. Before implementation, extend the packaged actual-handler
test to establish missing-entry/reopen RED. Preserve existing reader, draft,
save/version/conflict, expectation, guide-to-run and presentation GREEN identities.

Protect Close/reopen restoration, repeated-entry reuse, same-guide immutable append,
stable identities/source bytes, Cancel, older-version conflict with retained edits,
permission/current-policy/integrity loss, changed selection/context, parent/child
closure and four-size visible layouts. Denied readers retain their existing read
and Use access. No hidden-content deletion, automatic expectation rewrite, selected
run substitution, evaluation, business-Events publication or preference write is allowed. The
runtime files expected to change are Core guide draft/source/model boundaries and
Operations guide editor/reader ownership/forms; new focused helpers must stay
bounded. Build/compile all five packages, then relevant regressions, maintenance,
live-role, full-chain and operator evidence remain completion gates. This paragraph
records a compatible discovered control under D18 semantic inheritance, not
implementation or acceptance. The paired restart gate is currently independent;
do not alter its running test sources or frozen packages.

The first broad edit attempt stops at 235 PASS / one COM harness interruption,
before any new edit assertion. The first rejection is the read-only publication
counter (`0x80010001`), followed by cleanup `0x800AC472`; cause remains unproven.
Excel exits later without intervention, separately from the controller's initial
ExcelClosed=False. The focused actual-handler fixture then establishes **17 PASS /
34 expected FAIL**, all 51 unique typed assertions reached, without harness failure.
All 39 edit assertions run; twelve compile/fixture and five preservation/read checks
pass. Normal closure, zero matching Application events and unchanged sources/frozen
candidate are verified. Eight shared test helpers retain their exact original
function bodies. [Published-guide edit evidence](../../invSys_fork/tests/integration/plan022_slice4be_published_guide_edit_results.md)
records both attempts and exact commands/reports. Implement the specified entry,
captured published draft and immutable revision only after this recorded RED;
focused GREEN and all applicable completion gates remain required.

The first isolated implementation now passes **51/51** with all original focused
RED identities retained. Five packaged projects compile; normal Excel closure,
zero matching Application events and unchanged running sources/candidate are
verified. Six editor captures were directly reviewed across four sizes, reopen and
conflict. The normal draft notice is blank after restored fields load. Extend the
same packaged test before fixes to protect that notice, explicit disabled-edit
reasons for restricted/denied entry, and four reader sizes. This enforces the
existing D18 contract, not a new architecture. The 51 prior passes remain required.
Static evidence also identifies one new duplicated tag formatter (193 groups
versus 192); its two bodies/callers are reviewed for shared consolidation. Preserve
all other candidates. Final focused/visible GREEN, maintenance ratchets, broader
regressions, live roles and full chain remain open; first candidate is frozen at
`deploy/validation-published-guide-edit`. Exact evidence is in the linked edit record.

The notice extension now verifies **55 PASS / three expected FAIL**, including
all original 51 GREENs and four reader layouts, with normal closure and no matching
Application errors. Correct only the draft notice lost during field restoration
and missing Core edit reasons. Position the existing reader status below provenance
to retain its normal read notice and append the reason without shrinking minimum
supported form sizes. Consolidate the reviewed tag formatter and then build a new
isolated candidate; preserve the first candidate and this behavioral RED.

The corrected isolated candidate, `deploy/validation-published-guide-edit-notices`,
now passes **58/58**, retaining all notice-RED identities, with all five packages
built/compiled and ten individually reviewed reader/editor/reopen/conflict captures.
Normal closure, zero matching Application events and unchanged candidate/test
sources are verified. The reviewed Core formatter consolidation restores duplicate
groups to 192; dynamic calls remain 9/45 and all 28 prior module-size limits hold.
Compiled comparison identifies only the seven intended existing components and two
new helpers; 55 frozen package hashes are preserved. The broader visible gate must
retain the preceding 305 identities plus 46 edit assertions on this candidate.
Relevant role regressions, fresh-process preference, live roles, full chain and
human acceptance remain pending. This is a recoverable focused checkpoint, not
completion of Slice 4be or approval of either separate architecture proposal.

The first broader attempt records 286 console-observed preceding passes and one
COM harness interruption at the retry scenario's readonly form Count, before any
edit assertions. First-failure capture shows an unpainted owned view; cause remains
unproven. A subsequent temporary-fixture cleanup error prevents typed result and
controller-exit serialization. Preserve the console-derived evidence separately.
The exact owned process has a verified typed zero workbook count and closes by
assisted normal Quit; no force/recovery deletion occurs. Unchanged sources/candidate
and zero matching Application events are verified. Repair report retention with
offline coverage before retry; this changes developer tooling, not the runtime
contract. Any opt-in retry must be bounded to the exact readonly probe, preserve
first-failure evidence and retain single execution of action handlers. Focused
58/58 and its ten reviewed captures remain valid, but the broader gate stays open.

Offline tooling proof now passes 8/8 result-retention checks and 38/38 diagnostics/
retry/privacy checks, retaining the original ten no-action-replay identities. Typed
results are saved before generated-fixture cleanup and updated with explicit cleanup
failure. The opt-in retry covers only the exact readonly Action Path view Count
probe for the two observed COM rejections, with four total attempts and bounded
250/500/750 ms delays. Preserve the first failure and recovery trace; no action or
service call is replayed. Retry the complete gate against the unchanged frozen
candidate; do not claim the COM cause was resolved. Automatic review blocked manual
removal of the old generated temporary fixture, which remains in place. New gates
use fresh isolated fixtures; no recovery data or operational workbooks are changed.

The complete visible retry now passes **351/351** on the unchanged corrected
candidate: every prior 305 identity plus all 46 published-guide edit assertions.
Root `7eafdbb4722f47eabf0e5eacba52b07f` runs **2026-09-22 05:03:12.0300621--
05:35:35.0440291 UTC**. All 52 captures are individually reviewed/hash-verified.
Excel closes normally, matching Application events are zero, and running sources
and packages are unchanged. No first-call failure or readonly retry occurs; this
does not diagnose the previous COM cause. Regenerated static evidence preserves
246 components, 6,000 procedures, 132,040 lines, dynamic calls 9/45, 192 duplicate
groups and all 28 prior module-size limits. Role/restart/live/full-chain and human
acceptance remain separate gates. Neither pending architecture proposal is approved
by this evidence-only checkpoint; D13 does not require inventing a new product RED
for an acceptance-record update.

The corrected guide-edit candidate also passes the fresh-process gate **27/27**
at `ad220eef446e40bbabeae5885c443a2f`, **2026-09-22 05:36:26.4660945--
05:39:06.7599580 UTC**. Original owned Excel closure is normal, followed by a
different verified process with Operations/Core/Domain and no Admin dependency.
Saved Compare both restores over an unsaved How-To switch; no automatic pair or
evaluation is inferred. Exact explicit guide/run binding and all Config/training/
probe bytes are preserved. All three captures are directly reviewed and hashed;
final closure, unchanged sources/packages and zero matching Application events
are verified. Role/live/full-chain and human acceptance remain pending.

Corrected-candidate role regression now retains Viewer **94/94** at
`f55f414706fc4880b6ed974073c45470` and Boxing/Shipping **1,707 prior GREEN checks**
at `9c6cbd7005fe4502ae8fffccd1defe11`, with only the same seven missing-Auth recreation
failures under pending D8-A. All 3/22 captures are individually reviewed/hash-verified.
Both gates close Excel normally without intervention, preserve sources/packages,
and have zero matching Application events. Boxing assertions finish before a slow
normal process exit; the next gate waits for verified closure. No new runtime fix,
Auth decision or Event Detail scrolling acceptance is inferred. Diagnostic evaluation,
live roles, full chain and human acceptance remain required.

Diagnostic regression on the corrected candidate now passes **376/376**, retaining
every prior identity, at `decf781980bd431596c6208f0ee6a630` during **2026-09-22
06:01:10.1828270--06:21:20.8405962 UTC**. All 19 captures are individually reviewed:
pending four-source evidence, partial three-applied/one-awaiting, all-applied
conclusion and explicit expectation editing at supported sizes. Normal Excel
closure, unchanged sources/packages and zero matching Application events are verified.
The existing pane-activation helper retries once for Pending/Minimum before a
successful capture; this is not a product failure.

The same corrected guide-edit candidate subsequently passes full chain **32/32**,
live roles **48/48**, and Create Warehouse **15/15**, preserving all preceding
chain/live identities. Evidence prefix `reports/runtime/published-guide-edit-chain`
runs **2026-09-22 06:22:24.5090280--06:30:44.5998916 UTC**. Terminal exit is zero;
all three tracked reports are restored byte-for-byte, all 233 running-source pins
and 55 frozen-package pins are unchanged, and matching Application events are zero.
Cleanup requires assisted normal Quit after verified native process identity and
a fresh typed zero workbook count. The recovery dialog is directly reviewed;
**Yes, I want to view these files later** is selected and reviewed before confirming.
Excel exits and local test settings restore. No forced termination or recovery
deletion occurs; do not report unattended cleanup. The linked published-guide edit
record contains the verifier and exact evidence. Remaining D18 work and human
acceptance remain open. This acceptance-only update creates no runtime contract,
requires no invented D13 RED and approves neither pending architecture proposal.

**Direct tracked-action curation D13 entry:** The normative D18 refinement now
names Action Paths **Choose tracked actions**, its Operations-owned source picker,
exact loaded-publication binding and reuse of the existing guide editor. Core
validates selected original activity bodies and current maintenance/visibility
policy; guide schema 1 retains an empty SourceRun rather than fabricating a run.
This refines the already approved non-recording authoring route. It grants no
execution, source-write or automatic conclusion behavior.

Before runtime edits, add/run focused packaged tests using actual Admin Save
activity outside a recording, ordinary Events publication/loading and the new
picker's real handlers. Expected behavioral RED is missing entry/selection/draft/
save behavior. Protect multi-selection across search, exact original observations,
stable authored StepIds, empty SourceRun, None expectation, immutable save/revision,
current-policy/capability/context/source guards, parent/child lifecycle, reuse,
four layouts and source/config/training non-mutation for non-save actions. A broken
fixture, compile or harness is not RED. Runtime scope is Core draft/source/model
handling plus Operations library/picker/editor; preserve the current frozen
guide-edit candidate and all accepted GREEN identities. Focused GREEN, packaged
build/compile, visible evidence, static ratchets, relevant regressions, live roles
and full chain remain required. No deployment or full Slice 4be acceptance is
claimed by the refinement; Event Detail/D8-A approval remains pending separately.

Direct curation's focused packaged RED is now verified **12 PASS / 41 expected
FAIL**, 53 unique typed checks, at `afe5b11ae907444681cc750897d9322a`, during
**2026-09-22 06:43:54.8095241--06:46:07.0742576 UTC**. The preceding 49-check RED
retains every identity/outcome; four added parent/Viewer/sign-out/permission-loss
checks also fail for the missing runtime behavior. Actual unrecorded Admin
activity/publication, five instrumented compiles, startup and preservation checks
pass. Normal unassisted Excel closure, unchanged 175 test-source and 55 package
pins, and zero matching Application events are verified. Protecting helper and
exact command/evidence are in code
`tests/integration/plan022_slice4be_guide_action_curation_results.md`. Implement
only the specified Core/Operations path next; the same assertion set must pass
GREEN on a new isolated candidate, followed by all required supporting gates.

The first isolated curation candidate builds and compiles all five packages, but
source selection rejects valid published actions: activity publication lines are
27-field hashed envelopes, while the existing guide observation schema is the
26-field body. Preserve both contracts. D18 explicitly clarifies verification of
the existing envelope digest before copying its original body into the guide;
never mutate the loaded publication or add the hash field to guide observations.
The focused test's body comparison must likewise distinguish the original body
from its wrapper, checking the original activity-file digest and every retained
body value. Retain the first failed candidate/evidence; run the corrected protecting
test against it before fixing the source adapter. This is schema clarification,
not a new source/guide format or approval of an architectural exception.

The corrected direct-curation candidate now passes **55/55** protecting focused
checks and **63/63** expanded integration checks, including the actual published
reader/Edit/Save round trip. Stable step IDs, unchanged original bodies, empty
SourceRun and exact immutable predecessor linkage survive that round trip.
All five isolated packages build/compile; twelve library/picker/editor captures
are directly reviewed at four sizes. Normal unassisted closure, unchanged source/
package pins and zero matching Application events are verified. A preceding
read-only COM rejection remains separately recorded as a harness interruption,
not behavioral RED or acceptance. Maintenance retains dynamic calls 9/45,
duplicate groups 192 and all 28 previous size limits. Existing broader guide,
fresh-process, Viewer, Boxing/Shipping, evaluation, live-role and full-chain gates
are next on this same frozen candidate. Exact commands/results and limits are in
code `tests/integration/plan022_slice4be_guide_action_curation_results.md`.
This implements the approved D18 refinement; it does not approve pending D8-A or
Event Detail changes, transfer scope, deployment or human acceptance.

The first broader guide regression reaches 137 passing checks before Excel
rejects a read-only guide-form count with `0x800AC472`. Normal unassisted shutdown,
unchanged runtime/packages and zero matching Application events are verified;
the attempt is retained as a harness interruption, not accepted regression.
The separate opt-in `RetryGuideObservationForTest` handles only that exact count
and the previously rejected picker Values getter, with the existing bounded
four-attempt policy and first-failure retention. Offline test-first evidence is
70 PASS / 8 expected FAIL -> 78/78 GREEN, retaining all prior 38 identities and
single execution of actions. The full packaged gate is restarted on the same
frozen candidate; no runtime or architectural contract changes.

**Direct-curation source-integrity D13 follow-up:** Read-only review identifies
uncovered malformed RecordId and attempt/result CatalogVersion, PackageSetVersion
and BuildIdentity agreement cases. D18 already requires immutable valid activity
identities and agreement on the producing package/build/catalog; its direct-source
picker inherits those rules. The recorded-run reader already enforces that release
agreement. After the active regression closes, extend the actual Choose tracked
actions test with separately rehashed disposable publication fixtures for each
case, retaining one valid action alongside the changed group. Expected RED is the
invalid group remaining selectable; unrelated valid actions and original source
bytes must remain intact. Confirm behavior before changing private Core validation.
This is protection of the current contract, not a new source schema or approval
to rewrite evidence. No runtime defect is claimed solely from the code review.

Packaged source-integrity RED now confirms **67 PASS / four expected FAIL** over
71 unique typed checks, retaining all prior 63 identities. Separately rehashed
fixtures prove malformed result RecordId and mismatched catalog/package/build
remain selectable; each valid neighboring action and original source bytes are
preserved. A preceding 59-check fixture interruption is not behavioral RED.
The second broader guide attempt is separately interrupted at 286 passing checks
by a COM rejection while opening the recording library; no action is replayed.
Both completed attempts close Excel normally, with unchanged running sources/
packages and no matching Application errors. See the curation integration record
for exact roots/times. Tighten private Core validation under existing D18, then
build a new isolated candidate and require all 71 GREEN plus supporting gates.

**Activity-line label D13 entry:** Under D18's contributing-line label refinement,
the Operations picker will display each User activity line's published caption
and observed outcome, with a source-neutral prompt. Exact inventory key labels,
repeated lines, key accessors and cached read-only selection remain unchanged.
The focused actual Admin Settings Save -> ordinary publication -> Viewer list
selection test must fail for the current two Unavailable labels and inventory-only
prompt before editing Operations. Extend the existing Boxing published-read
checks for the same labels and all exact inventory keys. This is an inherited
presentation clarification, not a new identity, authority or execution contract.
Preserve every current GREEN, packaged compile/layout/static, live-role/full-chain
and visible evidence requirements; full guide/comparison acceptance stays open.

The focused unchanged-package RED completes 36 PASS / two expected label/prompt
failures. The Operations-only correction then builds all five isolated packages,
passes explicit compile/cold start and gives 38/38 focused GREEN with an inspected
Admin capture. All 225 compiled identities remain; only the detail controller and
form change. Static evidence retains 191 duplicate groups, 9/45 dynamic calls and
1158/1160 maintenance candidates; the display helper adds one method and 17 lines.
Detail passes 34/34, published Viewer/filter/Shipping state 94/94 and evaluation
376/376. Boxing/Shipping retains all preceding GREENs with 1707 PASS / seven known
D8-A FAIL; all 698 Boxing checks and 18 unchanged owner-return records pass.
Nineteen diagnostic/editor captures are inspected or byte-verified; activity and
business labels preserve original evidence. Ten current/frozen package hashes,
325 historical pins, protected sources, both unrelated user documents, all static
candidate identities and 28 growth limits pass.

The first full chain stops with an RPC harness exception (0x800706BE) during
projection rebuild after 32 live-role passes; Create Warehouse is 15/15. It is
not a passing chain. A verified empty residual Excel instance is closed normally;
the inspected recovery prompt retains files for later viewing. Settings/reports
are restored and the original controller terminates with Excel closed. The nine
audited gate windows contain no Excel Event1000, which does not explain or negate
the RPC failure. A fresh unchanged-candidate chain after confirmed cleanup passes
32/32, with live roles 48/48 and Create Warehouse 15/15. It also requires normal
closure of a verified empty Excel residual and retention of recovery files.
Settings/reports are restored and Excel closes; its audited window contains no
Application events 1000/1001/1002. Final preservation passes. This is an assisted
functional chain, not proof of unattended recovery or a native repair. The code
label evidence retains exact reports and scope. Continue 4be.5 authored guides and
actual How-To/Diagnostic/Compare handlers with this unchanged packaged baseline;
the entire Slice 4be and Release 1 remain unaccepted.

Activity-label checkpoint: code **9bfac38**, isolated candidate
`deploy/validation-event-detail-labels`; controls v1.158 records the same scope.

**Boxing published Viewer supplementary proof:** On the unchanged
`deploy/validation-boxing-activity-visible` candidate, actual Make/Unbox evidence
passes through the ordinary Admin publisher and actual Viewer/Detail handlers.
Test/evidence checkpoint: code **17e6677**; runtime remains **8ccb5b9**.
The expanded route completes 1689 PASS / seven known D8-A FAIL; all 112 added
publication checks and 680 Boxing checks pass. All preceding 1584 identities,
1577 GREENs and 18 owner-return records are retained. Six new captures are
inspected. Five instrumented compiles, unchanged fresh static metrics/component
sizes, all 28 growth limits, 325 historical package pins and the five candidate
hashes pass. Excel closes with no Event1000 in the audited run window. No runtime
contract, implementation or deployment changes; earlier full-chain/live-role
results retain their original scope and are not rerun by this supplementary test.

The captures expose an activity-picker usability gap: both observation lines
show **Unavailable** because they have no System_Key, under an inventory-only
prompt. Their selected detail fields and exact references are correct. Next
clarify fixed activity caption/outcome labels under D18 semantic inheritance in
the normative display rule and controls, then establish real-handler RED before
changing Operations. Never invent an inventory key or discard repeated lines.
The code evidence `plan022_slice4be_boxing_published_read_results.md` records
the exact report, capture limitations and unchanged broader requirements.
Boxing-specific policy/detail changes, remaining coverage, combined recordings,
guide lifecycle and How-To/Diagnostic/Compare, physical and human Release 1
acceptance remain open. D8-A is still unapproved; no new native repair is claimed.

**Boxing final candidate visible GREEN:** The isolated candidate completes
1577 PASS / 7 known D8-A FAIL across 1584 checks; all 568 Boxing checks pass.
Every preceding check and all 1553 GREENs remain, with all 18 independent owner
records unchanged. Sixteen captures are inspected, including actual Make/Unbox,
tracking notices and three saved-path layouts. Five-package compile/cold start,
225 compiled identities with exactly seven reviewed changes, package preservation,
fresh static evidence and all 28 module limits pass. Dynamic calls remain 9/45;
duplicate groups decrease by one. Final build/compile/test windows have no Excel
Event1000, without proving repair of the earlier native failure. The functional
Release 1 chain passes 32/32, live roles 48/48 and Create Warehouse 15/15, with local
settings and tracked reports restored. Its separate audit records one combase.dll
access violation consistent with the restart/cleanup interval; an exact failing
call is unproven. Those first-run functional passes alone did not establish a clean
native chain gate. Settings passes 191/191, retaining all 187 prior GREENs; four added checks
exercise both Boxing editors and saved flags. Ten Settings captures are inspected
and that native-event window is clean. Full Receiving passes 854/854, preserving
the recorded 845 baseline plus nine accepted native-focus/closed-workbook checks;
its native-event window is clean. Action Path evaluation passes 374/374 with every
prior GREEN retained, nineteen inspected or byte-verified captures and no Excel
Event1000 in that window. Separate worksheet suites retain 215/215 guard and 210/210
scenario GREENs, with clean native-event windows. The fresh unchanged-candidate
chain passes 32/32, with live roles 48/48 and Create Warehouse 15/15, and no Excel
Event1000 in its audited window. All jobs end, Excel closes, and settings/reports
are restored. Final package/source, module-limit, parser, link and unrelated-change
preservation pass. The isolated Make/Unbox observation checkpoint is complete;
the code evidence retains exact reports and the guard's transient cleanup snapshot.
Earlier native causes remain unresolved; this successful repeat is not a repair.
D8-A remains unapproved. Comprehensive remaining control coverage, combined recordings,
guide lifecycle, both presentations/comparison, physical and human Release 1 acceptance
remain open. Accepted deployment and operational workbooks are unchanged.
The code Boxing evidence records exact report paths and preserved wider scope.

Boxing implementation checkpoint: code **8ccb5b9**, committed and pushed;
isolated candidate `deploy/validation-boxing-activity-visible`. Controls v1.153
and the maintained code evidence record the same verified scope.

**Boxing observation implementation entry:** The 1560-check RED route now
protects the approved catalog-9 Make/Unbox contract. Implementation reuses the
Shipping form's captured-context and permission guard, receives exact write-entry
references and required-step failure facts from Boxing owners, and records only
actual handlers. Legacy report presentation remains separate from observations.
Equivalent list-copy/report helpers are extracted to preserve existing oversized
module limits. Core/Operations change; accepted deployment remains untouched.
Existing tests retain their checks while recognizing catalog 9 and the extracted
helper location. Focused GREEN, compiled isolated packages, layout/visible evidence,
static maintenance, live roles and full Release 1 chain remain required.

The first compiled implementation candidate passes **544/544 Boxing checks** in a
**1553 PASS / 7 known D8-A FAIL** route, preserving all 1560 preceding identities
and 1388 prior GREENs. Supplementary catalog/reference regression passes 101/101;
all 31 catalog-8 definitions remain. Exactly seven compiled components change;
the form/owner module shrink to 2909/22375, all 28 limits hold, and dynamic calls
remain 9/45. The initial build has a native VBE7 access violation before Operations
source import; it exits and Excel closes before a fresh-directory retry. Retry
build/compile/test complete with no Event1000 in their windows, without establishing
a native repair. The next candidate captures original exception messages before
fact calls and is running the full Boxing route with owned-window captures and
saved-path layout/read-preservation checks. That candidate's GREEN, inspected
evidence and remaining Release 1 gates are pending; the code evidence records the
exact first GREEN and retained native-failure limitations.

**Boxing optional-tracking RED, 2026-09-14:** Tooling `d6bdcdf` is committed and
pushed. The unchanged runtime `5e2c45a`
completes **1388 PASS / 172 expected FAIL**. All prior 1486 checks/1321 GREENs
and Shipping recording40/40 remain. Six actual Make/Unbox cases add 67 passes
and seven expected failures: four missing tracking-unavailable notices and three
unsupported catalog-9 policy assertions. One submission and exact package/component
Inventory application pass in every case. Config/policy bytes, prior activity,
captured binding, staging/custom values and unrelated workbook survive. Catalog-8
policy is valid and excludes Boxing; catalog-9 disabled policy must validate before
its absence of collection can count as accepted disabled behavior. All 18 earlier
owner-fact records remain retained. Five instrumented compiles, unchanged fresh
static/component metrics, 28 growth limits, 325 package pins/protected sources,
two script parses and 84 links pass. Excel closes normally with no Event1000 in
the verified window. No runtime, deployment, full-chain or human UAT change is
claimed. Next wire catalog-9 observations and pre-owner context rejection with
explicit owner submission/required-step facts, preserving all prior GREENs and
module limits; the code Boxing evidence records the exact report and verifier
array-count correction. D8-A remains separate and unapproved.

**Boxing optional-tracking D13 entry:** Continue 4be.1 through actual Make/Unbox
handlers with a physically unavailable activity store, saved catalog-8 policy,
and explicit catalog-9 disabled collection. Independent exact Inventory lines
and one submission protect the business action. Require fixed unavailable notices
for failed storage/older-policy exclusion, no implicit collection or policy repair,
and preserved Config bytes, prior activity, staging/custom values and binding.
Catalog-9 policy validity is expected RED on catalog-8 runtime; successful actions
under that rejected policy do not prove valid disabled-policy behavior. This adds
tests under approved D18, with no runtime or architectural change. Preserve all
1486 preceding checks and 1321 GREENs before implementing the observation wiring.

**Boxing submission/processing RED:** Tooling `029af12` is committed and pushed;
runtime remains `5e2c45a` without rebuild/deployment. The final corrected run completes **1321
PASS / 165 expected FAIL**, retaining all prior 1180 checks/1123 GREENs and Shipping
recording40/40. Nine cases through both actual handlers add 198 independent passes
and 108 missing-observation failures. All 18 redacted owner-fact records are retained.
Generic success persists after failed processing, staging and refresh; legacy sync
can remain True after a failed refresh. Pass explicit required-step facts from the
owner instead of interpreting a generic success flag or report in the observation.
This follows the existing D18 Boxing refinement and changes no architecture.
Five instrumented compiles, four script parses, 84 links, fresh static metrics and
all 28 limits pass; 325 package pins/protected sources and unrelated files remain.
Excel closes with no Event1000 in the verified final window. The initial compile
failure and subsequent source-array/metadata-retention calibrations are documented
in `plan022_slice4be_boxing_activity_results.md`; none is runtime repair or UAT.
Next finish actual-handler tracking-store/disabled/old-policy tests, then wire the
catalog-9 observations, explicit step facts and proven context guard within existing
module limits. The broader control/recording/guide/comparison and Release1 gates,
including the separate D8-A decision, remain open.

**Boxing submission/processing D13 entry:** Extend the actual Make/Unbox route
under the existing D18 Boxing refinement with deferred/failed processing,
failed staging/refresh, server refusal, lost/exceptional acknowledgment,
uncertain acceptance and allocated-before-write refusal. Reuse the independent
Core write-entry probes, installing them before all five instrumented compiles.
Snapshot direct submission counters before catch-up enters the inbox writer;
observe actual owner returns and exact applied component/package lines separately.
Missing outcome/reference observations are expected RED. Fixture/compile errors
are not RED. Preserve the full prior 1180 checks/1123 GREENs and existing package,
source, static and user-file safeguards. Tracking/policy coverage and all broader
acceptance remain open; no new architectural behavior is proposed.

**Boxing context/permission RED:** Tooling `b6de0e5` is committed and pushed;
runtime remains `5e2c45a` without rebuild/deployment. The expanded run finishes **1123 PASS / 57
expected FAIL**, preserving all prior 1074 checks and 1041 passes, including
Shipping recording40/40. The new matrix contributes 82 passes, 16 context failures
and eight missing denial-observation failures. Signed-out forms reach the service
but its existing permission check prevents mutation. Same-user new sessions and
another authorized warehouse allow the stale form to enter the mutation owner;
the probe stops before writes. Revoked permission correctly stops mutation but
has no REQUESTED/DENIED observations. D18 already prohibits the observed stale
dispatch; no new architecture decision is needed. Both authority files, staging/
unknown values, captured workbook and unrelated workbook are preserved; Auth
restoration/relaunch/reuse pass. Five instrumented compiles, 325 package pins and
protected sources plus all static limits hold. Excel closes with no Event1000
in the verified window, no duplicate or harness failure. This is test evidence,
not runtime completion or a native-crash repair. Next protect failed staging/
refresh, uncertain submissions and optional tracking/policy before owner wiring.

**Boxing context/permission D13 entry:** The committed initial RED checkpoint is code
`2593385` / docs `deea147`. Extend it through actual Make/Unbox handlers with
calibrated service/mutation entry probes for signed-out, same-user reauthenticated,
other-target and same-session revoked-SHIP_POST cases. The alternate warehouse
grants SHIP_POST to distinguish binding failure from permission refusal. The probe
stops only at mutation entry; authorization and form handlers stay real. Preserve
all preceding 1074 checks/1041 GREENs, authority bytes, staging/unknown values and
captured workbook. Missing context guards and denial observations are expected RED;
failed setup/compilation is not RED. No runtime or architectural contract changes.

**Boxing Make/Unbox D13 entry:** Continue Slice 4be.1 under D18's Boxing refinement.
Catalog 9 names BOXING_MAKE and BOXING_UNBOX with existing SHIP_POST eligibility;
runtime remains catalog 8 until meaningful packaged RED is established. The test
will use the actual form handlers for Make, Unbox and zero-quantity rejection,
independently observe submitted IDs/owning refresh returns, and check recorded
order, captured binding, exact keys and source-line preservation. Missing activity
and recording observations are expected RED; setup/compile failure is not RED.
Preserve the preceding 40 Shipping recording checks and comparable expanded
Shipping GREENs. Context/permission/failure matrices remain required before runtime
completion. This compatible control registration follows semantic inheritance;
the pending D8-A decision and all broader Release 1 gates remain separate.

The frozen-package run completes **1041 PASS / 33 expected FAIL**, preserving
every prior 1016 check and 1009 pass, including Shipping recording40/40. Boxing
has 32 independent owner/source/preservation passes and 26 missing-observation
failures; seven existing D8-A failures remain separate. No duplicate or harness
exception occurs. Five instrumented projects compile; Excel closes with no
Event1000 in the verified window. All 325 package pins/protected sources, runtime
static metrics/component sizes and 28 module limits remain unchanged; three
scripts parse and 84 links resolve. This is meaningful RED, not implementation
completion. Next protect captured-context/permission loss and failed staging/
refresh: the current refresh routine can return True with a staging warning, so
the normal-return probe cannot justify clean outcomes for that branch. Both
Shipping form/owner modules are at their size limits; retain those limits through
reviewed, protected extraction when implementing. The code repository's
`tests/integration/plan022_slice4be_boxing_activity_results.md` records exact scope.

**Shipping-only 4be.4 recording proof:** Tooling checkpoint `7c0307e` is committed
and pushed. The unchanged runtime source `5e2c45a` candidate
`deploy/validation-recording-notice` passes 40 focused assertions around the actual
Viewer Start/Stop controls and eight Shipping form actions. Evidence preserves
ordered occurrences, actor/warehouse, distinct repeated Add ActivityIds, every
submitted source ID and sixteen exact observation bodies in the integrity-linked
journal. An initial test compared the outer activity-file hash with the observation
body; the corrected body comparison passes without a runtime change. The complete
Shipping route retains its separate scope and pending D8-A findings. Two earlier
setup failures remain recorded; moving probe installation before fixtures and
compiling all five instrumented projects does not establish a native-crash repair.
Controls v1.148 and the code repository's
`tests/integration/plan022_slice4be_shipping_recording_results.md` record the scope.
No normative contract is changed by this D18 evidence addition.

The completed route records **1009 PASS / seven existing D8-A FAIL**, with every
one of the preceding 1016 check identities and 1008 passes retained. All 963
comparable earlier Shipping passes remain; four prepared-order fixture checks
belong to the alternate ordering. There is no harness exception or duplicate
identity. Five instrumented compiles, unchanged static metrics/28 module limits,
325 package pins/protected sources, three script parses and 84 links are verified.
Excel closes normally and neither completed calibrated run records Event1000.
This does not resolve earlier native failures, approve D8-A, deploy packages or
claim a new full-chain/human acceptance pass. Continue with actual Boxing
Make/Unbox observation and context/source-reference RED under D18, then the full
cross-role recording proof below.

**Remaining 4be.4 recording proof:** The current combined recording fixture exercises
Receiving and Admin; Shipping mentions in its reader/capture helpers only check that
reads do not access Shipping authority. It does not prove Production, Boxing and
Shipping participation in the shared recorded sequence. Extend the actual packaged
handler evidence for those roles under the existing D18 contract, retaining actor,
warehouse, separate captured workbooks, ordered occurrences and all source IDs.
The full-role business chain cannot substitute for this recording proof. This is
an evidence gap, not a demonstrated runtime defect or a new architecture decision.
Guide authoring, both presentations, comparison and import/export remain required.

Runtime/tooling checkpoint: code `44f0334`, isolated candidate
`deploy/validation-library-title`; controls v1.145. This checkpoint is limited to
the title correction and its verified evidence, preserving the full D18 objective.

The initial visual run stops at 109 PASS / one Admin-bootstrap harness exception,
before the new checks. A native Excel failure and WER-linked recovery instance are
recorded; the WER module-version-mismatch label does not establish root cause.
Both blocking recovery dialogs are resolved with retain-for-later selected and
visually verified. Excel exits without forced termination or recovery-file removal.
All 315 package pins/protected sources remain intact before the unchanged retry.

The retry repeats 109 PASS / one bootstrap harness failure. A Git-blob-verified
copy of the committed harness passes bootstrap but stops at 185 PASS / one
processor Excel/RPC failure; the new probe is not required for a native failure.
Scoped diagnosis with the viewport consolidated into the existing control probe
reaches 51 PASS / 6 FAIL: seven Pending layout/preservation checks pass, five native
window captures fail, and the processor loses Excel/RPC. The fixture is corrected
to set Excel visibility before creating Viewer/library. Its next run fails during
Admin Seed before reaching that correction (1 PASS / one harness failure).
No pane capture, native repair, current full-gate success or slice completion is
claimed. Preserve these failures and isolate fixture/processor native instability
before repeating the complete visual gate. The recording results document retains
exact commands, limitations and safe recovery evidence. Runtime remains unchanged.
Evidence-tooling checkpoint `cb9e080` retains this incomplete status. Final checks
preserve 315 package pins/protected sources, parse four scripts, resolve 83 links
and retain the existing maintenance limits. Excel is closed and all jobs terminal.

**Selection-continuity checkpoint (code `7c1bb5e`, pushed):** `validation-evaluation-selection` builds
and explicitly compiles five packages plus Operations cold start. Only the five
intended compiled components change. Its full packaged route passes **324/324**,
retaining all **293** preceding passing identities and adding **31**. All eighteen
binding checks, complete terminal-reference validation, verified-UTC display and
original library/editor layout and preservation gates pass. Core's declared primitive
`modPathEvaluation.SelectedBinding(context, pathId)` exposes only an opaque session
selection token; evidence objects stay in Core and no token is persisted as identity.

The four editor foreground captures pass and are visually inspected. Diagnostic-pane
capture and human acceptance remain separate. The initial build fails at Operations
workbook creation, then succeeds unchanged after Excel closes. The first full run
ends at 19 PASS / one native/RPC harness failure; Event1000 identifies an Excel access
violation with unknown module. The successful run temporarily stops responding,
then resumes on the same live handle without intervention and closes normally.
Neither observation proves a native repair. Static evidence preserves 192 duplicate
groups, 9/45 dynamic calls and 28 module limits. Release1 chain and limits/storage
verification were then run: the chain passes 32/32, live roles 48/48 and Create
Warehouse 15/15 with settings/tracked reports restored. Event1000 still records a
native Excel failure (`c0000409`, module unknown) during the chain; no manual recovery
intervention is needed, but native reliability is not established. The separate
storage/limits gate passes 78/78 and retains every previous 77 passing identity.
The action/size bounds, continued ordinary work and no-partial-file checks pass.
All jobs are terminal and Excel closes without intervention. Final verification
preserves 315 package pins and protected sources; eight scripts parse and 83 local
links resolve. Full Slice4be and Release1 acceptance stay open.

**Selected-run expectation checkpoint (code `ea2e2ae`, pushed):** Following `1db6057`, the
shared editor extends to D18's selected-run analysis contract. Headless Core keeps
the selected journal's exact RecordId/version/hash binding and separate staged
intent; Operations owns one editor across recording and analysis. Changing run or
context clears the draft. Captured expectation and This evaluation remain distinct,
and no analysis edit writes the original journal, publishes a guide, or emits
workflow activity. The library retains its 720 by 520 minimum. This is an
implementation of the existing D18 refinement, not an architecture change.

The expanded actual-handler RED completes at 189 PASS / 60 FAIL, retaining every
prior 189 passing identity without a terminal harness exception. Its two new
failures require nonempty draft invalidation on run switch and prevent inherited
steps in a new selection. The shared editor/Core binding implementation follows
that RED. The isolated candidate `deploy/validation-analysis-expectation` builds and
explicitly compiles all five packages, including Operations cold start. Comparison
finds exactly seven changed components and the two new owner modules. Its packaged
run completes at 210 PASS / 39 FAIL, retaining all 189 prior passing identities and
gaining 21 GREEN checks. Analysis opening/default/cancel/edit/stage/provenance and
run-switch behavior pass, as do prior recording/editor and library-layout checks.
The sign-out result test is tightened to require a populated conclusion before
sign-out, preventing an absent control from passing as cleared evidence. It remains
RED with the evaluator. The remaining failures are 35 evaluator checks and four
foreground captures. Static growth is 195 VBA lines, with all 28 existing module
limits, 192 duplicates and 9/45 literal/unresolved calls preserved. Excel closes
normally. The same candidate passes full chain 32/32, live roles 48/48 and Create
Warehouse 15/15, including verified local-settings restoration. Tracked reports
are restored exactly. Final checks preserve 290 package pins and protected sources,
parse 17 changed scripts and resolve 83 local links. No result/visible acceptance
or native-reliability repair is claimed.
The complete evaluator, immutable derived records, guide lifecycle, both
presentations and Release 1 acceptance remain required.

**Captured editor/writer implementation checkpoint (code `1db6057`, pushed):**
The captured editor/writer implementation has been validated against the
175/62 RED checkpoint within the scope recorded below. Operations owns the modeless editor and Core owns a
separate headless draft bound to the originating context/SequenceId. Primitive
`modActionRecording` editor calls return EXPECTATION1 projections; fixed draft
edits never dispatch the registered workflow. Add generates a distinct StepId,
reorder preserves it, and Use validates/stages the explicit terminal definition.
Stop freezes it into schema-2 Close and discards the draft; Start/Observation
retain None. Sixteen source-import harnesses include the new draft service.
Controls v1.138 records the editor and minimum layout. The final isolated candidate
`deploy/validation-captured-expectation-retry` builds/compiles all five packages,
including Operations cold start. The full packaged gate reports 189 PASS / 58 FAIL,
retaining all 175 previous passing identities with no duplicates or terminal
harness exception. Writer/captured staging, compatibility, caption, three layouts,
exact reorder, stale binding and sign-out pass. Two additional actual-editor checks
first prove the incorrect False retry default and its persisted effect, then pass
after initializing the add-step choice True as D18 already requires. Explicit False
still persists after reorder and Stop. No architectural change or repeat approval
is required for this correction.

The remaining failures are 54 pending evaluator checks and four visible-capture
failures. Read-only diagnosis finds no foreground window and input-desktop access
denied; visible operator evidence remains pending. Static regeneration retains all
28 existing large-module limits, 192 duplicate groups and 9/45 literal/unresolved
calls. The caption candidate passes full chain 31/31, live roles 48/48 and Create
Warehouse 15/15. The final retry candidate passes action-limit/storage checks
77/77. The first controller-exit run reports 120 passing checks then three harness
failures before proving a new explicit run. It also exposes a test-isolation
defect: the coordinator exits while a child Excel remains and loses its local-
settings restoration copy. The fixture is deselected without guessing the user's
prior warehouse. The coordinator now retains restoration state while waiting;
sanitized worker diagnostics exclude exception text and arguments. Lifecycle/
redaction and private-pipe tests each pass 6/6. The diagnostic restart rerun passes
123/123, closes both controllers/Excel and restores its starting deselected state.
The full-chain validator's temporary selection is also observed and removed only
for its exact completed fixture. It now shares the private in-memory settings
snapshot/restore helper, waits for Excel closure, and verifies values/types.
Lifecycle/settings/redaction checks pass 10/10 and private-pipe checks 6/6.
The final candidate passes the strengthened full chain 32/32, retaining every
prior 31 check and adding LocalSettingsRestored; live roles 48/48 and Create
Warehouse 15/15 pass. Tracked reports are restored exactly, Excel closes, and the
known deselected starting state is restored. All 285 package pins/protected
sources are preserved, 26 scripts parse and 83 links resolve. This tooling repair
is not an architectural change, native-reliability fix or recovery of the original
local selection lost by the first failed coordinator. Evaluation, visible evidence
and full Slice 4be/Release 1 acceptance remain open.
The recording results document records the Document Recovery cleanup obstruction
and verified retain-for-later closure. It does not explain earlier native failures.
This implements existing D18; evaluation, guides, both presentations, comprehensive
coverage and full release acceptance remain open.

**New-writer test entry (code `8728e08`, preceding implementation):**
The new-writer gate checks actual original journal entries separately from reader
format variants: schema 2 throughout, None on Start/Observation, and default None
on ordinary Stop. Existing lifecycle/identity facts cover both supported formats;
new writing specifically requires schema 2. This protects the existing D18 writer
contract before its implementation. Review also adds `modExpectationModel` to all
16 source-import lists that import `modRecordingModel`; those declarative harness
repairs parse and add exactly one existing dependency each. They do not alter the
packaged candidate or establish new product behavior. Controls v1.137 records the
same test entry; editor/captured staging and full release scope remain open.

The unchanged candidate completes at 175 PASS / 62 FAIL, retaining all 175 prior
passing identities without a harness exception. The three added writer checks
are the only added RED; all 265 package pins and protected sources remain intact,
and Excel closes normally. No additional runtime/build/full-chain acceptance is
claimed. Next implement the captured editor, exact-sequence staging and schema-2
writer using these RED cases and the existing editor-binding tests.

**4be.4 recording test entry (historical):**
The next test-first entry is the D18 saved-run read surface: actual Viewer
Action Paths, library selection/Search/Refresh, complete journal validation,
missing Close, corrupt/missing links and current policy. The frozen recording
foundation supplies real Start/Save Value/Stop journals. Absent product controls
are expected RED; missing callable instrumentation remains a harness failure.
Core owns the read projection; Operations owns one captured library instance.
The frozen foundation records71/18 RED across89 identities, with all68 prior
checks retained GREEN and no harness exception. The first reader candidate
compiles all five packages and passes89/89, with ten foundation/reader package
pins preserved and Excel closed. Static duplicates189->193 remain unresolved;
the copied builder and short event bodies require cleanup/review before the
maintenance gate passes. Reader layout, expanded integrity/context cases,
actual Excel restart, broader regressions and visible acceptance remain open.

Reader follow-up tests enforce matching journal/observation package-set and
catalog provenance, owning attempt/result build identity, release warnings, warehouse/schema rejection and
library geometry/context recovery. Builder duplication is removed through a
declarative control table. Explicit maintenance exception RDR-UI-THUNKS-01
retains only the three single-statement MSForms duplicate groups named in the
recording evidence report (Hide, guarded Layout, and RefreshPaths delegates).
Expected duplicate ceiling192 preserves the189 baseline plus those three
required event bodies; no other growth exception is granted. D18 architecture,
D13 tests and all remaining reader/release acceptance gates remain unchanged.

The expanded frozen-reader run records99/5 across104 checks: two common
package-set/catalog failures are valid RED; the proposed Core/role build
equality is excluded by the correction below. Two proposed older-build/package
warnings use an unsupported chronological assumption. BuildIdentity is a GUID
and PackageSetVersion is a compatibility label. Per the normative clarification,
those two cases must instead require a difference warning with age unavailable;
the existing lower-catalog Older release rule remains. Rerun corrected warning
expectations before implementing that presentation; preserve all89 prior GREEN
identities and original-evidence retention checks.

The first combined candidate exposes a correction to the proposed build test:
Activity.MakeBody obtains BuildIdentity from its producing Admin/Operations
package, whereas the journal records Core. Equating those IDs breaks accepted
recording and contradicts D18 owner provenance. Remove that comparison, retain
the common package-set/catalog checks, and test mismatched owning attempt/result
builds instead. This restores the normative contract; it is not permission to
rewrite original activity or change package build identities. The failing
candidate cannot supply reader or Release1 GREEN evidence.

The rebuilt owner-provenance candidate compiles all five packages and passes
117/117 focused checks and59/59 filters, preserving prior77/89 identities and
all104 expanded IDs with the documented expectation corrections. Static audit
confirms192 duplicate groups with exactly the three explicit exceptions, no
dynamic-call growth and all28 individual oversized-module limits. Publication
stops before assertions with the recurring RPC setup failure; its verified
empty recovery child closes with three add-in hashes unchanged. The unchanged
candidate then passes publication82/82 and detail34/34. Full-chain validation
passes31/31, with live48/48 and Create Warehouse15/15, preserved candidate hashes,
restored generated reports and Excel closed. The chain records another
combase/c0000005 Excel fault despite passing assertions. Native-clean, actual
recording restart and user acceptance remain unproven.

The recorded-run foundation now has35/29 RED ->64/64 packaged GREEN on the
first compiled candidate. The expanded256-action baseline is36/32 RED across68
identities. Core journal/session services and Operations Start/Stop/Cancel/status
controls implement this foundation, with explicit incomplete outcomes at known
context/policy boundaries and failed persistence. The cleaned candidate compiles
five packages and passes77/77 extended limits/storage/layout checks, retaining
all prior50/64/68 identities. Full chain31/31, live48/48, Create Warehouse15/15,
filters59/59 and Settings187/187 pass. After two publication harness/native
failures, serial unchanged-baseline and recording-candidate runs each pass82/82
with normal Excel closure; no publication code change was made. A native
Excel fault during the passing chain and unavailable foreground/input desktop
keep clean native and visible acceptance open.
No full recording, journal-reader, comparison or Release1 acceptance is claimed.

The normative D18 journal refinement specifies Start/Observation/Close entries,
generated path/version names, distinct record/run identities, hash links,
atomic 1MiB-bounded writes and original activity bodies. Missing Close is
Interrupted; reads never resume a recorder. Add real-handler checks for durable
Start, incremental observations, immutable Close, policy/context interruption
and denied storage before implementing Core recording. These are refinements
of the approved non-authoritative recording contract, not new task authority.

`Test-Slice4beConfigCommands.ps1
-CheckActionRecording` retains the 30 published-Viewer checks, then probes the
actual Start/Stop/Cancel controls and existing Admin Save Value handler. Missing
product controls are explicit observations; the callable instrumentation must
work and the ordinary attempt/result fixture must be valid before recording RED
can be claimed. Protect disabled capture, sequential occurrence identity,
incremental immutable activity, stopping without a conclusion and cancellation.
The normative recording status is `lblRecordingStatus`. This first lifecycle
test does not establish all-role coverage, saved-record integrity, limits,
restart/interruption or diagnostic evaluation; those remain required 4be.4 gates.
The unchanged compiled maintenance candidate records **34 PASS / 16 FAIL**,
50 unique checks, no harness exception, with all 30 prior published-reader
identities retained GREEN. Recording controls/status are absent; valid Admin
attempt/result pairs retain blank sequence IDs and zero ordinals. Four new
passing checks prove ordinary activity and byte preservation, not working
recording transitions. Five candidate hashes remain unchanged and Excel closes;
the run's Application Error 1000 window contains no Excel fault. Prior native
faults remain unresolved. That baseline claimed no recording implementation or GREEN.
See `tests/integration/plan022_slice4be_action_recording_results.md` in the code
repository. The subsequent durable-record and interruption/limit tests precede
the foundation implementation described above. Next, protect the full-journal
current-policy reader, missing/corrupt links, missing-Close restart behavior and
real cross-form Operations sequences before implementing those remaining paths.

Use a focused packaged harness such as the proposed
`tools/validate_slice4be_event_viewer.ps1`, with fixtures entering through Admin
Generate Warehouse/Create Warehouse and optional Seed. A missing fixture,
missing callable test seam or compile failure is not behavioral RED; instrument
actual handlers if needed while preserving their operator path, then assert the
missing behavior. Direct service checks are supplemental. No new RED/GREEN is
claimed by this documentation pass.

**Acceptance cases that must remain explicit:**

- At least one normal Operations task and one Admin task, including multi-event
  submission and later processor application, appear through the shared Viewer.
  Every reachable eligible handler has a coverage assertion, not just sample
  happy paths. Excluded pre-sign-in/creation-without-runtime actions and
  unavailable old Admin history have visible reasons.
- How-To and Diagnostic use the same path/version and selected observed run.
  Authored/reordered steps never alter the original sequence. Compare both
  identifies matched/missing/extra/failed/unavailable steps. Imported examples,
  disabled tracking, missing source result, stale publication and partial capture
  cannot become a successful local conclusion.
- Recording off/on, command versus navigation collection, current policy
  visibility (including saved guides), actor/context/capability checks,
  mid-sequence policy changes, 256-action limit, 1 MiB rejection, NAS failure,
  duplicate deliveries, cancellation and Excel interruption are protected.
- Viewer preserves repeated System_Key detail lines, different keys in one
  event, unlike UOMs, source timestamp provenance, current-state supplements,
  actual Hold versus internal reservation, accepted date filters/list export,
  and byte-for-byte non-mutation of inspected source authority.
- Settings reset stages changes only; whole policy/profile saves reject stale
  versions, unknown fields and dirty/locked/missing-required Config. Existing
  scalar Settings/UOM behavior and D5 read non-mutation remain GREEN.
- Preserve the existing technical baseline: 18/18 D5, 81/81 packaged smoke,
  48/48 live-role, 30/30 full chain, five explicit compiles, Viewer/layout,
  all-role launcher and reusable Production restart checks. Add new coverage;
  do not reduce or relabel those counts to hide regressions. Keep Excel
  validators serial and check cold-start package references.
- Regenerate static maintenance; no unexplained growth in bloat/dynamic-call
  metrics, no scanner-driven deletion. Review source/package hashes, layout
  geometry and visible forms at default/minimum/maximize/restore.
- Visible operator evidence: edit tracking/profile on the dedicated tab, choose
  each view and Compare both, record an Operations task and an Admin task,
  author/reopen a guide, inspect actual completion and an incomplete run, and
  reopen saved workbooks without lost binding or duplicate launchers. Fresh
  human acceptance remains distinct from automated evidence.

**Expected file/package scope:** Operations `frmInventoryViewer.frm`,
`modInventoryViewer.bas`, an Operations-owned personal Settings surface,
Admin `frmAdminSettings.frm` and its controller, affected real Operations/Admin
handlers, Core projection/publication modules and dedicated D5 commands,
bounded activity/Action Path services, tests and five-package manifests.
Core/Domain remain headless. Domain changes need focused RED for a necessary
owning read/publication boundary; no canonical business schema change is
proposed. Runtime reports are redacted/ignored; no operational workbook writes
or NAS rollout follow automatically from product-contract approval.

**Historical evidence at review preparation (before approval):** The 2026-09-07 source read confirms Settings
currently has one runtime canvas and no Event Tracking tab; current Admin audit
entries use a station workbook with no stable event ID. Prior Viewer inspection
found Inventory/Events/ListBox->Table and no Action Path/profile handlers. The
new proposal changes no code, XLAM, active architectural contract or acceptance
status. Validate document consistency, links/diff and Git status now; D13 begins
with behavioral RED after contract approval.

**Historical 2026-09-06 source-readiness evidence (before approval):**

| Boundary | Verified source | Consequence for the first packaged tests |
|---|---|---|
| Event publication | Core `modWarehouseSync.WriteSnapshotEventRows` copies twelve named fields from `tblInventoryLog` and deletes/recreates target rows. | Preserve unknown columns and every event detail line; add source/coverage evidence without inventing missing history. |
| Event identity | Inventory Domain `modInventoryApply.ApplyEvent` writes the same EventID/AppliedSeq for every line in `linesToApply`. | An EventID is not a unique log-line key. Multi-output and repeated-key lines must survive grouping and detail selection. |
| Viewer read envelope | Core `modInventoryViewerData.LoadCurrentInventoryEventViewerData` emits ten display fields, drops EventID/System_Key, and leaves the UOM field blank. | Add protected identity/detail data; test all exact keys and UOMs through actual Viewer selection, not only service output. |
| Shipping supplements | `modTS_Shipments.AppendBoxDesignViewerEvents` opens ShippingBOM read-only; `AppendHeldShipmentViewerEvents` reads current hold TSV and uses its file timestamp. | These are current-state sources, not durable Hold/Design event history. Moving reads to publication must preserve visible coverage without claiming historical timestamps/actions. |
| Designs history | Designs Domain `modDesignsApply` writes `tblDesignEvents` with EventID. | Use its owning read/publication boundary; current inventory-event publication does not include it. |
| Admin activity | `modAdminConsole.AppendAuditEntry` writes `tblAdminAudit` in the resolved Admin workbook, with no stable event ID. | Do not fabricate an ID from worksheet position or assume a central publisher can discover every station's log. Coverage/correlation for this source remains a design gap to resolve before its implementation. |
| Admin settings | Before the approved D5 correction, `frmAdminSettings.mBtnSaveConfig_Click` called the mutating `modConfig.UpdateConfigValue` reader API. | The approved correction routes the handler to headless `Core.modConfigCommands`, keeps the old scalar entry as forwarding compatibility only, and makes ordinary reads non-mutating. Later profile/capture commands use this ownership; their D18 schema/model remains proposed. |
| Time | Inventory `ApplyEvent` and Admin `AppendAuditEntry` write `Now` into UTC-named fields. | Column names do not prove UTC. The proposal now labels unverified historical zones and requires verified UTC for new publication/load metadata. |

Real action seams include Viewer `OpenInventoryViewer`,
`RunInventoryViewerEventsForTest` -> `frmInventoryViewer.TestEventsReport` ->
`mBtnRefresh_Click`; Receiving `mBtnAdd_Click`/`mBtnConfirm_Click`; Production
`mBtnProcessSave_Click`/`mBtnProcessRelease_Click`,
`mBtnRecipeSave_Click`/`mBtnRecipeRelease_Click`,
`mBtnManagerCheckIn_Click`/`mBtnManagerApplyOutput_Click` (Complete Run); Shipping
`mBtnHold_Click`/`mBtnRemove_Click`/`mBtnSend_Click`; and Boxing
`mBtnBoxBuilderSave_Click`/`mBtnBoxMakerMake_Click`/`mBtnBoxMakerUnmake_Click`.
The capture allowlist must distinguish genuine operator invocations from
programmatic handler calls; the source names alone do not prove this distinction.

Validation run against code `0afe05f`: Slice 4bc source assertions **5/5 PASS**
and Slice 4w source assertions **12/12 PASS**. The regenerated Slice 4w report
now correctly names all three Viewer tabs. Read-only inspection verified all
five `deploy/current` XLAM names, sizes, SHA-256 hashes and package-set values
against `addins-manifest.json`; this does not prove source/build correspondence.
The checked-in `inventory_viewer_results.md` still describes two tabs although
the current harness asserts three. Its prior PASS is historical, not a fresh
packaged baseline. The harness directly constructs Config/Auth/inventory fixtures
rather than entering through Admin Generate Warehouse; its fixture setup must
be assessed against D14 before treating a rerun as Release 1 acceptance. No COM
run, new behavioral RED, compile, layout, deployment, or UAT occurred in this
readiness pass.

#### Slice 4be historical proposal (superseded by the 2026-09-03 amendment)

This is a proposal only. It must not change the existing read-only **Viewer**,
its bounded R1 **Events** projection, its list export, or any event payload
until Architecture v4.11 and the controls catalog are explicitly approved.

The proposed comprehensive Viewer would remain a projection: the processor
would publish a bounded, cross-domain event snapshot for the selected
warehouse, and Viewer would read that published snapshot only on explicit
Refresh. It would never open a canonical event/inventory/design workbook for
writing, refresh authority, process an inbox, or use a global Aggregator
snapshot as warehouse authority. Candidate standard columns are readable local
time, event family/action, reference/correlation, affected item/entity,
quantity/UOM, location/condition, source role, and user. Exact `System_Key`
remains available only as the immutable entity identity, never as a row
surrogate. Filters, row limits/paging, freshness, and retention must be
specified before implementation.

The proposed Admin detail-profile model is a warehouse-scoped versioned Config
projection. An Administrator could select only whitelisted event-detail fields
for each event family and set their display order/labels. The profile changes
future Viewer rendering only; it never rewrites historical event payloads,
unhides secret/credential data, or makes an operator projection authoritative.
The required Admin surface, profile schema, default profile, capability gate,
and cross-XLAM primitive-envelope boundary remain unapproved design work.

The user requested an optional **Action Path** to speed training by revealing
the user-facing controls used to perform a task. Approval must choose exactly
one of these non-hybrid models:

1. **Durable captured path:** role handlers append a bounded sequence of
   approved user-facing control captions/identifiers to the event at creation;
   Viewer optionally reveals that recorded sequence. It is audit history, not
   hidden implementation telemetry.
2. **Derived training path:** Viewer maps each published event family to a
   versioned list of user-facing controls; no per-event control sequence is
   written. It is training guidance, not an audit assertion that every mapped
   control was clicked.

Both models exclude backend processor calls, hidden controls, credentials,
workbook paths, and `Application.Run` mechanics. A later D13 slice must begin
with public packaged Viewer/Admin/role-handler RED proving read-only authority,
detail-profile capability/configuration boundaries, Action Path correctness,
filtering/freshness, exact-key visibility, and no regressions to current
Events/list export. No implementation is authorized by this proposal.

**2026-09-03 approved amendment — this replaces the conflicting proposal
language above.**

**Slice 4bd** is now **multi-server Aggregator source selection and physical
two-warehouse proof: implementation in progress; isolated D13 GREEN; physical
two-computer UAT pending**. Architecture D17
supersedes the single-current-server restriction without changing warehouse
authority. `invSys.Admin.xlam` will add an `ADMIN_MAINT` aggregation source-set
form reachable from **Aggregate Global Snapshot**. It discovers warehouses
readable through the current server connection, lets the user add another
authenticated NAS/server connection for that session, and selects published
snapshots to aggregate. It visibly explains selected/skipped/rejected sources
and leaves the normal Send To target unchanged.

The public action copies/reads only selected published snapshots, validates
source identity/freshness/compatibility, rejects ambiguous duplicate
WarehouseIds, and writes an advisory output to the designated existing output
feed. It never opens source authority for writing, creates a warehouse, or
stores server credentials. One-source aggregation remains the default when the
form is not used. D13 begins with packaged Admin form/action RED for current
and additional-server discovery, selection, duplicate/rejection explanation,
source-authority non-mutation, and exact WarehouseId/System_Key preservation.
Only then may physical UAT use two distinct NAS-backed roots and two
Windows/Excel stations.

**2026-09-03 implementation record.** The public Admin callback now opens
`frmAggregationSources`. The session-only form discovers currently connected
and remembered NAS roots on open; the user does not re-enter credentials for an
existing Windows connection. **Add Server** exposes Windows credentials only
when an additional server is needed. A
source is selectable only when its readable runtime `tblWarehouseConfig`, its
configured published snapshot, and the snapshot WarehouseId agree. The form
shows WarehouseId, server root, snapshot path, freshness, source fingerprint,
READY state, selected-source count, and rejected/skipped state. It does not
alter Send To. The Admin-to-Core call serializes the selected file list as a
primitive envelope; Core reads only those files and records their source
identity in the advisory output. D13 evidence: source-set contract 6/6 GREEN,
updated Admin callback contract 3/3 GREEN, isolated Core/Admin build GREEN,
explicit two-source `WarehouseId`/`System_Key` aggregation GREEN, and packaged
form initialization GREEN. Physical two-server/two-computer UAT remains
required before Slice 4bd acceptance.

**Historical 2026-09-03 Slice 4be record, superseded by current D18:** **comprehensive Viewer and curated Action Path: approved;
implementation pending**. Architecture D18 approves a versioned curated
training path, not either former captured-control or derived-control model.
Viewer remains read-only for inventory/event authority. An `ACTION_PATH_MAINT`
user selects currently visible Events in intended order, supplies
name/tags/instructions, and saves a non-authoritative Action Path library
record. Any signed-in Viewer user can search/read published paths. It is never
executable automation or a claim that selected events occurred at an import
target.

Action Path shows origin, version, selected-event availability, and a
non-blocking release/schema age warning. Export/import uses a hashed,
versioned JSON training package with no credentials, workstation paths, or
authority data; import assigns a new local identity and preserves origin so a
fake warehouse can rehearse instructions safely. Admin detail profiles remain
read-only rendering configuration and require a whitelisted, versioned Config
contract. D13 begins with public Viewer/Admin handler RED for event
selection/save, search/version warning, export/import validation, capability
denial, unchanged authority, current Events/list-export regressions, and
package restart. Visible acceptance requires create/search/staleness/import
evidence.

**Slice 4bh — Admin Viewer-event projection control, archive-first retention,
and save-feedback diagnosis: approved; implementation pending.** Architecture
D19/D20 lock the default-on `AdminViewerEventLoggingEnabled` setting as a
Viewer-noise control only: it cannot suppress canonical audit/security or
authority records. A separate `ADMIN_MAINT` Data Lifecycle surface begins with
inspection and hashed archive creation for explicitly non-authoritative
collected data; automatic/destructive retention is disabled. Save notices are
diagnosed by owner before invSys-owned routine feedback is changed; invSys never
changes Windows/NAS-client/global Office notification policy.

**Slice 4bi — non-Admin first-use NAS onboarding: approved; implementation in progress.** A new Windows/Office profile with the five XLAMs installed can have no remembered NAS root, which previously made Operations **Server Sign In** fail with Admin-only setup wording even if the user's Windows SMB session already had access to the warehouse server. The approved D-NAS correction permits the existing Core-owned connection form only after an explicit Operations **Server Sign In** click when no root is remembered (or a saved root has unusable Windows credentials). The user enters an authorized UNC root, scans using current Windows access or explicitly connects with NAS credentials, and selects a validated target. The result persists through the existing per-profile Core root/target APIs only. It never grants an invSys role, creates/repairs configuration, changes another target, or folds Windows/NAS identity into invSys sign-in. D13 begins with the public `ToggleServerSessionForCapability` RED proving this first-use fallback, non-Admin wording, separate target/user layers, and absence of credential persistence; GREEN requires focused source/public-handler validation, package compilation, static maintenance, and the applicable Release 1 chain.

**Slice 4bj -- NAS self-service station setup: implemented and visibly accepted.** A user must not obtain an individual XLAM from GitHub to start invSys. The NAS deployment feed now publishes a stable **StationSetup** entry point that resolves the current immutable release, validates all five packages against its manifest, installs them side-by-side in the user's local cache, and registers the Operations/Admin startup leaves only when Excel is closed. It may offer periodic-updater registration but works without it. The entry point is read-only against warehouse data and requires no existing warehouse target or invSys identity; it uses NAS access only to read publisher-controlled deployment files. The post-install path remains Server Sign In -> target selection -> invSys Sign In. Create New Warehouse remains visible/reachable only after a real `ADMIN_MAINT` sign-in and is additionally constrained by current Windows/NAS write permission; setup access is not an authorization grant. D13 RED was 2/6, then 6/6 source GREEN and 9/9 isolated NAS-to-cache/leaf-registration GREEN; the pre-existing D16 suite remains 18/18 GREEN. Physical feed proof on 2026-09-04 ran the ordinary NAS `.cmd` launcher and applied `R1-20260904-977d2a2` to the current station; a second user then successfully logged into `jwj-nas-1` and installed the same command launcher.

**Slice 4bl -- Explorer-compatible NAS revalidation: implemented and visibly accepted.** A user can have a valid current Windows SMB session in File Explorer yet see invSys reject a remembered root after `WNetAddConnection2` reports Windows error 67. Core validates the already reachable UNC folder through its filesystem provider before treating WNet as a reconnect fallback. The observed entry path can also contain redundant leading backslashes, so Core canonicalizes those to the two-separator UNC form before validation; it does not accept a protocol prefix or nonexistent root. The reported UAT exposed a separate form contract defect: **Scan** required the user to supply a root, rather than discovering a candidate. Server Sign In now exposes **Scan Roots**, which enumerates the current Windows user's visible SMB server/share candidates without opening them. The operator selects a root, supplies Windows/NAS credentials, uses **Connect**, and only then does Core scan and list the root's validated warehouse runtimes. If no candidate is discoverable, typed root remains an explicit fallback. It does not bypass NAS ACLs, persist credentials, create/change a target, or grant a role. D13 RED was 5/8; GREEN is 8/8. Local Windows proof confirms the current NAS deployment connection exposes `invSysWH1` as a visible Disk share through the same read-only discovery path. Five-package compilation and Ribbon regression passed. Static maintenance records the justified scoped increase from 5,399 to 5,406 procedures: five Core discovery helpers plus two separate form-action handlers; the sole additional dynamic root is the convention-dispatched `mLstRoots_Change` event protected by the focused handler test. On 2026-09-04, visible operator UAT confirmed NAS sign-in and the subsequent separate invSys sign-in both passed.

The feed-root `Install-invSys-Station.cmd` wrapper is the ordinary-user entry
point. It starts the publisher-controlled NAS bootstrap with a process-scoped
PowerShell execution-policy bypass so Windows does not reject a network-share
`.ps1` before it can perform the existing release/hash checks. It contains no
credential or authority logic and cannot substitute NAS access for invSys
authentication or `ADMIN_MAINT`.

**Slice 4bk -- generated user onboarding packet: implemented; visible message review pending.** The User & Roles form creates or updates warehouse authorization but does not send email. Its existing clipboard action is now **Copy Account & Setup** and generates the human-deliverable station handoff: the stable NAS `Install-invSys-Station.cmd` entry point, authorized NAS/Tailscale precondition, selected warehouse scope, and install -> Server Sign In -> target -> invSys Sign In sequence. It does not provision Windows/NAS access, include NAS credentials, select a target, or claim installation grants invSys authorization. D13 RED was 2/6; GREEN is 7/7, including the actual `mBtnCopyPin_Click` form action. Five-package XLAM compile passed, ribbon regression is 48/48 GREEN, station-setup regression is 9/9 GREEN, D16 deployment regression is 18/18 GREEN, and regenerated static evidence shows the edited form added no procedure or dynamic-call root. The full COM packaged validator stalled without a result and was terminated; it is not counted as passing evidence.

### Slice 4bf -- clean SharePoint/GitHub/NAS deployment, automatic update, and rollback: implemented; physical deployment UAT passed

This approved slice implements Architecture v4.11 D16 without changing the
five-package/headless package boundary or warehouse/domain authority. No XLAM
may be built, copied, replaced, registered, or rolled back while any Excel
process is open.

The release model is an immutable five-package release directory under
`<PathSharePointRoot>\Addins\Releases`, with `current-release.json` in the
parent Addins root. A manifest records package name, package-set version,
SHA-256, build commit, build time, and compatibility metadata. Publication
verifies a complete new release before atomically changing the small pointer;
it retains the current release plus two earlier verified releases. GitHub
remains source/review authority; an optional NAS mirror preserves that exact
release layout but is never inside a warehouse runtime root. No package release
contains warehouse data, auth/config, inbox/outbox, user credentials, or
operator workbook state.

The station updater is a Windows Task Scheduler job that runs at user logon and
every 15 minutes, but only applies an update when Excel is closed. It verifies
the selected release manifest/hashes, copies the complete set into a
station-local versioned cache, retains the known-good complete set, and then
changes only the account-scoped Operations/Admin leaf startup registration.
At task installation, the small station-maintenance PowerShell toolset is
verified and staged under the local invSys deployment directory; the scheduled
task targets that local copy rather than a Git checkout. Git remains source/
review authority only. This is not an XLAM, separate launcher, or part of the
five-package release.
It fails closed and preserves the active known-good set if a package, manifest,
hash, write, or registration check fails. It never updates a loaded XLAM in
place, never edits an authority workbook, and emits only redacted local
diagnostics. The successful applied version is visible after the next invSys
session starts; no interactive update prompt interrupts a user.

An update failure after staging automatically restores the last known-good
release. A deliberate rollback is limited to a local Windows administrator,
again only with Excel closed: it selects one retained previously verified
release, validates its full manifest, repoints the station to that complete
set, and records the reason/time/version locally. Rollback is a package/runtime
rollback only: it never rolls back processed inventory events, warehouse
snapshots, designs, config/auth data, or an operator workbook.

The approved trigger is automatic application before the next Excel/invSys
session without an operator prompt; the updater defers while Excel is open. It
retains three complete releases, needs no independent bootstrap launcher, uses
local Windows-administrator authorization for deliberate rollback, and uses the
Architecture D16 SharePoint/NAS layout. D13 begins with RED/GREEN for
incomplete-release rejection, hash mismatch, Excel-open deferral, full
five-package update, automatic restore, manual rollback, Operations/Admin
registration order, and canonical-workbook hash non-mutation.

**Implemented interfaces.** `tools/publish_invsys_release.ps1` publishes an
immutable release from a built `deploy/current` five-package set to the chosen
SharePoint/NAS Addins feed using an explicit `ReleaseId`. It writes and verifies
`Releases/<ReleaseId>/release-manifest.json` before atomically advancing
`current-release.json`. `tools/update_invsys_station.ps1` is the unattended
station action; it hash-verifies before side-by-side caching and invokes
`tools/register_current_addins.ps1` only for the Operations/Admin leaf pair.
`tools/register_invsys_update_task.ps1` displays the proposed task by default
and creates a verified local station-tool copy plus the logon/15-minute task
only with `-Apply`. A local Windows
administrator uses `tools/rollback_invsys_station_release.ps1 -ReleaseId
<retained-id> -ReasonCode <approved-code> -ConfirmRollback` with Excel closed.
The reason code is constrained to a non-sensitive approved value. No script
reads, writes, copies, or registers an inventory/design/configuration/auth or
operator workbook.

**D13 evidence (2026-09-03):** `Test-Slice4bfDeployment.ps1` was RED at 0/9
before the tools existed, then GREEN at 17/17. Its isolated feed proves
immutable manifest publication, Excel-open deferral, five-package hash
verification, Operations/Admin-only registration, third-party add-in
preservation, protected authority-workbook hash non-mutation, non-admin
rollback refusal, tampered-release rejection, and prior leaf-registration/
known-good-pointer restoration after injected registration failure. It also
proves that the scheduled-task preview targets a local verified station agent,
not a Git checkout. The existing five-package Slice 13 cutover regression
remains GREEN at 14/14.

**Physical deployment evidence (2026-09-03):** the current `ad53c52`
five-package build was published as immutable release `R1-20260903-ad53c52` to
the separate NAS deployment share's D16 Addins feed; its canonical warehouse
runtime root was not used. With Excel closed, this station hash-verified the
feed, cached the release, and repointed only Operations/Admin startup entries.
The local verified station agent was staged. The user then installed the
logon/15-minute `invSys.StationUpdate` task from an elevated session. Read-only
verification proved it enabled with both triggers and an existing local agent
updater, an on-demand scheduled run returned `0x00000000` with `APPLIED`
status for the release, and a clean Excel startup loaded both cached Operations
and Admin add-ins. The verification Excel instance was closed afterward.

### Slice 4bg -- historical inventory pattern/reorder worksheet workbench: proposed; awaiting explicit approval

This is a proposal only. It must not alter the existing **Create Inventory
Table** / **Upload Selected Inventory Table** contract, which remains the
Admin catalog/event path for current managed inventory.

The proposed historical workbench would add a distinct local worksheet/table
family in the captured saved Admin workbook for pasted historical transactions
and non-authoritative pattern analysis. Its required typed columns would be
historical date, item code or source item text, UOM, signed quantity/direction,
and source/reference; location, supplier/customer, category, cost, and notes
would be optional. The table must visibly say **Historical Analysis -- Not
Current Inventory**, never contain `System_Key`, never become an invSys managed
inventory table, and never be selected by the existing Upload Selected
Inventory Table action.

The proposed Analyze action would write only local formulas/projections in the
captured workbook: bounded date-range summaries, receipts/issues by source
item/UOM/location, period demand/usage, lead-time/reorder inputs when present,
and advisory reorder-pattern output. It must retain unlike UOMs as separate
groups, tolerate unknown end-user columns, preserve pasted source data, and
identify missing/invalid rows without mutating them. Its results are training
and planning aids; they do not create catalog items, allocate stock, create
events, affect on-hand/availability, update a warehouse snapshot, or call a
processor.

Approval must choose whether a history row may resolve an existing catalog item
for display-only comparison, which period/frequency defaults are appropriate,
the exact advisory reorder fields/formula, retention of local tables, and
whether this surface belongs in the existing Add/Edit Inventory Items form or
as a separately permission-gated Admin worksheet action. A later D13 slice
must begin with public captured-workbook RED/GREEN proving table separation,
no `System_Key`, no event/processor/canonical change, unknown-column
preservation, mixed-UOM separation, analysis correctness, and safe handling of
invalid history rows before any user-visible workbench implementation.

### Remaining Release 1 work recorded at the user checkpoint

The user identified the following work as required before Release 1 acceptance.
These entries preserve priority but do not silently define new architecture:

- [ ] complete visible acceptance of the existing Architecture v4.11 central
  Aggregator with two real warehouses operating on two computers. Automated
  packaged proof is GREEN: two isolated warehouse roots publish separately,
  the advisory global snapshot retains `WarehouseId` plus the exact
  `System_Key` (never collapsing same-SKU entities), and a later receipt at one
  warehouse catches up from 8 to 11 while the other warehouse remains 5. This
  proves the package contract, not operation on two physical computers;
- [ ] reconcile a comprehensive Viewer contract with Architecture v4.11,
  including Admin-configurable detail and an Action Path view that can explain
  operator workflow and user activity;
- [ ] clean and prove the SharePoint, GitHub, and NAS setup for end users, then
  define an easy-deploy/easy-update and automatic-updater contract before
  implementation; and
- [ ] reconcile an Add/Edit Inventory Items worksheet-import contract for
  historical data used in pattern/reorder analysis, modelled on the Process
  worksheet workbench without treating imported history as current managed
  inventory.

The Aggregator item advances existing normative requirements. The Viewer,
updater, and historical-analysis import require an approved normative contract
and focused implementation slices before runtime changes.

## 6. Batched user acceptance checkpoint

Request one user checkpoint only after Slice 4 automated evidence is GREEN.

Exact steps:

1. Close all Excel windows when requested and wait for confirmation that the
   new five-package set is installed.
2. Open Excel normally, use Operations **Server Sign In**, and verify the
   dedicated NAS test warehouse is selected.
3. Use **invSys Sign In** and authenticate to the selected warehouse.
4. Close any role operator workbook, then click **Receiving**.
5. Confirm one saved station-local Receiving operator workbook opens and the
   Receiving form opens modelessly.
6. Close only the Receiving form, click **Receiving** again, and confirm one
   form reopens against the same workbook without an automation error.
7. Click **Production** and confirm its station-local workbook and modeless
   form open without a type mismatch or missing-workbook instruction.
8. Click **Shipping** and confirm its station-local workbook and modeless form
   open without a type mismatch or missing-workbook instruction.
9. Activate a different ordinary workbook and confirm each open role form
   remains bound to its original operator workbook.
10. Use **Server Sign Out** and confirm both session controls return to their
    Sign In labels and role controls are disabled. Click **invSys Sign In** and
    confirm it instructs you to use Server Sign In. Then close and reopen Excel,
    reconnect/sign in, and repeat steps 4-8.
11. On the Admin tab, use **Send To** to select the dedicated NAS test
    warehouse and confirm the Operations selector/status updates to the same
    target. Click **Demo Inventory**, confirm Station shows this computer's
    Windows name, select **R1 Workflow Kit (built-in)**, and click **Seed Demo
    Inventory**.
12. Confirm the success dialog appears. Open **Inventory Viewer**, click
    **Refresh**, enter `DEMO-` in Search, and confirm the complete 24-entity kit
    is visible, including shipping carton, divider, label, tape, and void fill.
    Repeat the same built-in Seed and confirm it reports no newly created active
    groups. Use **Upload Data Set** to import a validation CSV, reopen the form,
    and confirm it is selectable. Seed it, then use **Delete Data Set** and
    confirm only the uploaded definition disappears: its seeded inventory
    remains. Confirm the R1 kit cannot be deleted.
13. Open Receiving, click **Refresh**, search for an item in **Receive item
    search**, select it from the dedicated results list, enter required
    Location and optional Lot, stage it, and confirm the top list remains
    **Receiving Entries History** rather than a duplicate inventory viewer.
14. On Returns, select one Condition-specific row and record its available
    quantity. Choose `RETURN`, enter a smaller positive quantity and a reason,
    stage and confirm it, then Refresh and verify both Returns and Inventory
    Viewer decreased by that amount without changing location, lot, or
    Condition. Repeat with `DUMP` on remaining inventory and confirm the same
    depletion behavior. Attempt an overdraw and confirm it is rejected.
15. In Process Designer, confirm Process/Requirement/Output IDs are generated
    locked three-character Base-36 values. Use **Send Process to Sheet** to enter
    100 lb sugar, 200 lb flour, 11.2 lb baking powder, and 300 lb filtered water;
    confirm the 611.2 lb batch basis and 100.0% formula total, then use
    **Retrieve Selected Process** and confirm the temporary table is removed.
    Send the Process back to the sheet once more to prove repeat editing. Create/
    release the Process with at least two outputs. In Recipe Designer, reuse it
    with another released Process, connect one output downstream, leave one as
    co-product, and confirm unresolved and circular edits are rejected. In
    Ingredients Assignment, map each external requirement to acceptable managed
    SKU alternatives. On Production Run - List, allocate exact inventory keys,
    complete two batches, and prove scales at the `0.001%`, `100%`, and `1000%`
    bounds. Production Run - Tree is not part of this checkpoint.
16. Open Shipping, resize **Box Designer** and **Box Maker** through grow,
    shrink, maximize, and restore; confirm full-width lists, aligned headers,
    non-overlapping actions, `NA` for items without a box alternative, and no
    zero-balance or same-`System_Key` duplicate Component inventory choices.
17. Use the seeded materials through Box Designer/Box Maker and complete one
    shipment. Confirm Shipping **Add** reserves the selected box without a
    `ROW`-column error, each public action reports staged, queued, applied, or
    refreshed state accurately, and record whether native Saving notices remain.
    After the five-box checkpoint, Shipping must show NAS Inv 95, Projected Inv
    95, and Locked 0 without relying on a later timer refresh.

Expected results:

- Receiving creates or reuses exactly one station-local saved operator
  workbook and opens its modeless form.
- Production and Shipping each create, open, or reuse exactly one station-local
  saved operator workbook and open one captured modeless form; neither returns
  `Type mismatch` or a missing-workbook instruction.
- No config, auth, inventory, snapshot, inbox, outbox, or XLAM workbook becomes
  visible as an operator workbook.
- Repeated clicks and restart do not create duplicate forms, workbooks, tabs,
  add-ins, or callback execution.
- Seed Demo Inventory completes against the selected dedicated test warehouse
  without flashing/hanging, an application/object-defined error, or Admin
  sheets/tables appearing in Config/Auth/inventory workbooks.
- The expanded seed result is not accepted until all 24 entities, including
  the five shipping-material rows, appear after Viewer refresh.
- Production content expands and restores with the native form window. The
  five top-level pages are Process Designer, Recipe Designer, Ingredients
  Assignment, Production Run - List, and experimental Production Run - Tree.
- Each completed Process output has a distinct new `System_Key`; routed
  intermediate output is consumed by that key and unconnected output remains
  visible as finished/co-product inventory.

Evidence to return:

- the full text of any dialog;
- whether each form opened;
- the operator workbook filename only, not its full user path;
- whether the second launch reused the same workbook/form; and
- the full Seed Demo Inventory result dialog and whether all 24 demo entities,
  including the five shipping-material rows, appeared after refresh; and
- screenshots of the Operations ribbon and each successful form, with
  sensitive warehouse/user data redacted.

## 7. Completion definition

This corrective plan is complete only when:

- [x] the read-only deployed runtime evidence identifies the actually loaded
  package hashes and selected NAS test root;
- [x] all three packaged launcher REDs are recorded for the observed behavioral
  reasons;
- [x] Receiving self-provisions or opens its station-local workbook;
- [x] Production and Shipping self-provision/reuse station-local workbooks and
  have no launcher type mismatch;
- [x] D15 Process Designer, Recipe Designer, Ingredients Assignment, and
  multi-output Production Run - List contracts are packaged GREEN through the
  actual operator handlers;
- [ ] visible saved-workbook/NAS Production acceptance proves reusable
  Processes, graph validation, exact-key allocation, multi-output creation,
  intermediate routing, co-products, persistence, and Events visibility;
- [x] plan 020 role, packaging, restart, static, and bloat regressions remain
  GREEN;
- [x] the dedicated NAS test-runtime checkpoint passes after clean Excel
  restart; and
- [ ] the user returns the batched acceptance evidence.
