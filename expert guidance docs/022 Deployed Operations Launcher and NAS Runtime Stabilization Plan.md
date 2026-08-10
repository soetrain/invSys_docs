# Deployed Operations Launcher and NAS Runtime Stabilization Plan

**Status:** Current corrective implementation plan
**Created:** 2026-07-28
**Applies to:** invSys Architecture v4.11, especially D2, D3, D12-D14, the
operator deployment model, and Phase 6 saved-workbook acceptance
**Scope:** Packaged Operations ribbon launchers, station-local role workbook
provisioning, deployed-package diagnostics, and dedicated NAS test-runtime
acceptance

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
- [ ] the operator confirms the visible Admin **Seed Demo Inventory** control
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
- [ ] repeat the visible dedicated-NAS seed/refresh checkpoint.

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

- [ ] add focused packaged form-action RED that proves the visible Shipping
  selection/staging path still exposes or depends on `ROW`;
- [ ] replace the form and backing path with exact `System_Key` identity;
- [ ] retain Shipping/Boxing lock, stage, hold/return, Shipments Sent, restart,
  and version regressions; and
- [ ] regenerate static maintenance evidence with no prohibited managed
  `ROW` control/header.

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
- [x] repeated seeding remains an intentional new-entity action, not an
  idempotent SKU upsert.

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

## 6. Batched user acceptance checkpoint

Request one user checkpoint only after Slice 4 automated evidence is GREEN.

Exact steps:

1. Close all Excel windows when requested and wait for confirmation that the
   new five-package set is installed.
2. Open Excel normally, use Operations **Connect Server**, and verify the
   dedicated NAS test warehouse is selected.
3. Sign in to invSys.
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
10. Close and reopen Excel, reconnect/sign in, and repeat steps 4-8.
11. On the Admin tab, click **Seed Demo Inventory**, keep the selected
    dedicated NAS test warehouse/station, and click **OK**.
12. Confirm the success dialog appears. Open the Receiving form, leave its
    **Receiving** tab selected, click **Refresh**, enter `DEMO-` in **Search
    inventory**, and confirm three additional demo entity rows are visible in
    the top **Inventory** list. Do not use Production's **Recipe Builder** page
    or Shipping's shippable-box list for this check: neither is the raw
    inventory view governed by this acceptance step.
13. Maximize and restore the Production form and confirm the four pages and
    their controls resize with the window rather than remaining at base size.

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
- The seed result is not accepted until the three entities appear after
  snapshot/operator refresh.
- Production content expands and restores with the native form window.

Evidence to return:

- the full text of any dialog;
- whether each form opened;
- the operator workbook filename only, not its full user path;
- whether the second launch reused the same workbook/form; and
- the full Seed Demo Inventory result dialog and whether the three demo items
  appeared after refresh; and
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
- [x] plan 020 role, packaging, restart, static, and bloat regressions remain
  GREEN;
- [x] the dedicated NAS test-runtime checkpoint passes after clean Excel
  restart; and
- [ ] the user returns the batched acceptance evidence.
