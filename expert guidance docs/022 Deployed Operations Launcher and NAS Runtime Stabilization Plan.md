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
    locked three-character Base-36 values. Use **Edit Process on Sheet** to enter
    100 lb sugar, 200 lb flour, 11.2 lb baking powder, and 300 lb filtered water;
    confirm the 611.2 lb batch basis and 100.0% formula total, then use
    **Retrieve Process from Sheet** and confirm the temporary table is removed.
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
