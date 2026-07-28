# Operations Stabilization and Developer Tooling Slice Plan

**Status:** Complete — Release 1 accepted
**Applies to:** invSys Architecture v4.11, especially D3, D9-D13, and Phase 6
**Scope:** Receiving, Production, Boxing, Shipping, Operations packaging, code-bloat control, and generated implementation/runtime evidence

## 1. Purpose and authority

This plan replaces the sequencing proposed in draft 019. It does not replace the v4.11 architecture specification.

Precedence is:

1. `invSys-Design-v4.11.md` is normative.
2. D12 defines the Operations package boundary.
3. D13 defines required test-first development order.
4. This document sequences implementation work under those decisions.
5. Generated manifests and runtime reports describe the implementation; they never override the spec.

The objective is not merely to make the current Production form pass one manual run. The objective is to:

- [x] stabilize Receiving, Production, Boxing, and Shipping as one coherent event-driven Operations system;
- [x] remove legacy direct-mutation and fallback paths that conflict with the Domain architecture;
- [x] reduce code bloat without deleting dynamic VBA entry points accidentally;
- [x] extract current implementation and runtime state on demand as JSON and generated Markdown;
- [x] consolidate the three role packages into `invSys.Operations.xlam` without creating another monolith; and
- [x] prove the complete Receiving -> Production -> Boxing -> Shipping chain through the packaged operator actions.

### 1.1 Implementation progress

Last verified: 2026-07-28.

| Gate | State | Code commit and evidence |
|---|---|---|
| [x] Entry gate | Complete | Authority pointers, branches, pull requests, and baseline state verified |
| [x] Slice 0 | Complete | `8d9f018` |
| [x] Slice 1 | Complete | `b9d472b` |
| [x] Slice 2 | Complete | `8776ee8` |
| [x] Slice 3 | Complete | `1b2f880` |
| [x] Slice 4 | Complete | `00e54a4`; `tests/integration/slice4_results.md` records focused RED/GREEN and the 15/15, 19/19, 59/59, 62/62, and 19/19 validation runs |
| [x] Slice 5 | Complete | `886f18b`; `tests/integration/slice5_results.md` records the 7/13 pre-refactor lock shape, packaged live layer isolation, saved Shipping 2/2, packaged XLAM 59/59, tooling 62/62, and maintenance baseline 19/19 |
| [x] Slice 6 | Complete | `ee0a376`; `tests/integration/slice6_results.md` records the 10/10 shadow contract, 13/13 packaged shadow validation, zero collision groups, legacy package 59/59, tooling 62/62, and maintenance baseline 19/19 |
| [x] Slice 7 | Complete | `eaed0f6`; `tests/integration/slice7_results.md` records focused 9/9, packaged two-batch PASS, shadow 13/13, packaged XLAM 59/59, tooling 62/62, maintenance 19/19, and restart-safe typed session evidence |
| [x] Slice 8 | Complete | `bb1111c`; `tests/integration/slice8_results.md` records focused 14/14, source Production 41/41, packaged two-batch PASS, shadow 13/13 with zero collisions, packaged XLAM 59/59, tooling 62/62, maintenance 19/19, deployed fixture stripping, captured-workbook authority, and primitive/JSON bridge evidence |
| [x] Slice 9 | Complete | `146c73a`; `tests/integration/slice9_results.md` records focused 8/8, packaged 3-size x 4-page and maximized geometry PASS, native window transitions PASS, inspected screenshots, source Production 45/45, shadow 13/13, packaged XLAM 59/59, tooling 62/62, maintenance 19/19, and the explicit declarative-layout diagnostic bloat exception |
| [x] Slice 10 | Complete | `51356c6`; `tests/integration/slice10_results.md` records focused 10/10, source Receiving 20/20 across workflow/surface/saved-runtime ranges, packaged Receiving action PASS, event replay proof, shadow 13/13, packaged XLAM 59/59, tooling 62/62, maintenance 19/19, and no static-ratchet regression |
| [x] Slice 11 | Complete | `29a80fe`; `tests/integration/slice11_results.md` records focused 11/11, exact `System_Key` release/restart/replay proof, packaged XLAM 60/60, live role workflow 46/46, RibbonX 156/156, shadow 13/13 with zero collisions, tooling 62/62, maintenance 19/19, and the reviewed tab-composition bloat exception |
| [x] Slice 12 | Complete | `cab79a9`; `tests/integration/slice12_results.md` records focused 11/11, protected Boxing behavior 3/3, packaged XLAM 60/60, live role workflow 46/46, RibbonX 156/156, shadow 13/13 with zero collisions, tooling 62/62, maintenance 19/19, five fewer runtime components, 164 fewer procedures, 68 fewer maintenance candidates, and 12 fewer duplicate-body groups |
| [x] Slice 13 | Complete | `b3d7363`; `tests/integration/slice13_results.md` records focused 14/14, exactly five deployed XLAMs with a hash-verified `R1-5` manifest, packaged compile/surface/restart 54/54, one-tab RibbonX 136/136, live consolidated workflow 46/46, shadow 13/13 with zero collisions, tooling 62/62, maintenance 19/19, typed same-project Ribbon callbacks, and legacy-role coexistence remediation |
| [x] Slice 14 | Complete | `39da706`; `tests/integration/slice14_results.md` records meaningful focused and behavioral RED, focused GREEN 9/9, the ordered packaged full chain 30/30, exact identity/balance/location assertions, restart and replay reconciliation, five-package runtime extraction, and the static retired-path ratchet |

Release 1 and plan 020's Entry Gate and Slices 0-14 are complete. The code
branch, including Slices 5-14, is pushed to
`origin/codex/fix-tester-station-nas-setup` on `soetrain/invSys_fork`;
documentation continues on the existing documentation branch and pull request.

Checked items in Section 7 are evidence-backed completion records. The
underlying architectural and regression requirements remain binding.

## 2. Current baseline and why tooling comes first

The role source currently has approximately:

| Role | Lines | Procedures | Literal `Application.Run` sites |
|---|---:|---:|---:|
| Receiving | 4,120 | 202 | 1 |
| Production | 17,937 | 724 | 15 |
| Shipping | 26,996 | 982 | 0 |

Largest files include:

- `src/Shipping/Modules/modTS_Shipments.bas`: approximately 21,625 lines
- `src/Production/Modules/mProduction.bas`: approximately 12,927 lines
- `src/Production/Forms/frmProduction.frm`: approximately 4,452 lines

Shipping is currently larger than Production and must not be hidden inside one late combined migration slice. Receiving is smaller but still requires its own contract and packaged form-action proof.

The existing deployment manifest records copied-file metadata. It is not a source contract manifest and not a runtime-state report. Two additional tools are required.

## 3. Tool A: static VBA maintenance scanner

### 3.1 Purpose

Create a deterministic, read-only scanner over exported `.bas`, `.cls`, `.frm`, `.frx` metadata, RibbonX, build-project definitions, and test sources.

Suggested entry point:

```text
tools/inventory-vba-surface.ps1
```

### 3.2 Required JSON model

The scanner produces versioned JSON first. Markdown is rendered from that JSON and is never edited independently.

Required sections:

- package/project composition;
- modules, forms, classes, and source paths;
- procedure names, visibility, kind, start/end line, and size;
- public and bridge entry points;
- direct calls and literal `Application.Run` targets;
- unresolved/dynamic `Application.Run` expressions;
- RibbonX tabs, groups, controls, callbacks, and capability gates;
- `Auto_Open`, workbook events, UserForm events, worksheet events, and processor handlers;
- table names, expected headers, and code locations that read/write/resize them;
- config keys, event types, capability names, and bridge contract versions;
- form controls, control types, page/frame ownership, and geometry when extractable;
- tests that reference each public entry point;
- duplicate normalized procedure-body candidates;
- same-project late-binding candidates;
- oversized module/procedure candidates;
- likely test fixtures or diagnostic code embedded in runtime modules;
- reachability confidence and reasons.

### 3.3 Dynamic-root registry

VBA reachability cannot be inferred from direct calls alone. Maintain a small committed root/exception registry for:

- Ribbon callbacks;
- `Auto_Open` and workbook events;
- UserForm/worksheet event naming;
- processor event handlers;
- string-dispatched macros;
- cross-XLAM bridge APIs;
- test harness entry points;
- COM or Windows callback procedures; and
- explicitly retained compatibility shims.

The scanner must never label a procedure safe to delete merely because no direct call was found.

Candidate confidence levels:

```text
HIGH    no direct/dynamic/root reference and no public contract
MEDIUM  no direct reference but public or convention-dispatched
LOW     dynamic call, event naming, callback, or unresolved contract
```

Deletion remains a reviewed source change followed by compile and regression tests. The scanner does not delete code automatically.

### 3.4 Bloat-control ratchet

Initial metrics establish a baseline rather than failing the build immediately.

After baseline:

- no module already over the agreed threshold may grow without an explicit exception;
- new runtime modules and procedures must stay under agreed size thresholds;
- same-project `Application.Run` counts may not increase;
- unresolved dynamic-call counts may not increase;
- new duplicate-body candidates may not increase; and
- test/fixture helpers may not be added to runtime modules.

Each cleanup slice reports:

- lines removed;
- procedures removed or moved;
- duplicate implementations consolidated;
- dynamic roots added to the registry;
- focused and regression tests run; and
- behavior intentionally preserved.

## 4. Tool B: read-only runtime-state extractor

### 4.1 Purpose

Static source inspection cannot answer which add-ins, workbooks, tables, versions, snapshots, or queues are active in the operator's Excel session. Build a separate read-only runtime extractor.

Suggested entry point:

```text
tools/export-invsys-runtime-state.ps1
```

It may use Windows PowerShell, Excel COM, and safe diagnostic APIs exposed by the XLAMs. It must not mutate, save, repair, refresh, process, or close operational workbooks unless an explicit non-default action says so.

### 4.2 Required runtime JSON

Capture:

- timestamp, machine/session identity suitable for diagnostics, and tool/schema version;
- loaded invSys XLAM names, full paths, hashes, project/contract versions, and `IsAddin`;
- duplicate or legacy role add-ins loaded beside Operations;
- open workbook names, paths, read-only state, visibility, and intended authority class;
- worksheet, ListObject, header, row-count, and table-range metadata;
- warehouse/station/runtime-root resolution and connection status;
- non-secret config keys and source scope;
- current invSys user identifier and capability names when safe;
- Inventory and Designs Domain bridge diagnostics;
- inbox row counts grouped by event type and status;
- processor lock/backlog/last-run metadata;
- snapshot/read-model freshness metadata;
- operator staging row counts;
- form/control inventory when a form diagnostic API is available; and
- warnings for workbook pollution, stale snapshots, duplicate identities, missing tables, or version drift.

### 4.3 Redaction and safety

Never emit:

- NAS/server passwords;
- GitHub/service credentials;
- PIN values or hashes;
- Windows credential material;
- secret tokens;
- unrestricted row-level inventory/customer data by default.

Row values require an explicit opt-in diagnostic mode and a documented redaction policy. Default reports use schemas, counts, identifiers, statuses, and hashes.

### 4.4 Generated output policy

Generate:

```text
runtime-state.json
runtime-state.md
implementation-manifest.json
implementation-manifest.md
maintenance-candidates.json
maintenance-candidates.md
```

Rules:

- JSON is canonical for generated evidence.
- Markdown is deterministically generated from JSON.
- Tool schemas and synthetic fixtures are committed.
- Machine/runtime reports are ignored by Git by default and attached to a test run or issue only when needed.
- A comparison command should diff two reports without requiring Excel to be reopened.

## 5. Shared workflow contracts: do not invent VBA inheritance

VBA class modules do not provide normal class inheritance. Do not design a `RunSession` base class hierarchy that the language cannot express cleanly.

Use composition and shared result/envelope contracts:

- role-specific `CProductionRunSession`;
- role-specific Receiving workflow/session state;
- role-specific Shipping/Boxing workflow/session state;
- a shared structured operation result;
- a shared event-submission result;
- shared state-transition validation helpers where behavior is truly identical; and
- JSON or primitive bridge contracts at cross-XLAM boundaries.

Receiving, Production, and Shipping do not have identical state machines. Standardize result shapes and transition rules, not business states that are only superficially similar.

### 5.1 Greenfield `System_Key` and extensible-header contract

`System_Key` is the immutable, system-wide unique identifier carried with each
durable entity wherever Excel displays or processes it. `ITEM_CODE`/SKU
identifies what an inventory item is; `System_Key` identifies the exact entity.
Location, quantity, `Condition`, and user-defined fields are attributes, not
identity.

Rules:

- new runtime tables use the exact managed header `System_Key`;
- `ROW` is deleted from the new runtime contract and is not renamed, copied, or
  migrated into `System_Key`;
- invSys does not import or repair legacy business inventory for this reset;
- new keys are generated only by the owning creation/service boundary and are
  globally unique across warehouses, stations, workbooks, and entity types;
- keys are immutable and survive sorting, filtering, refresh, save/reopen,
  movement, condition changes, and projection rebuild;
- events and relationships reference the affected `System_Key`;
- every table defines a required managed-header subset and tolerates additional
  end-user columns;
- refresh, resize, and rebuild logic preserves unknown/local columns and does
  not use ordinal column positions;
- shared custom values persist by `System_Key`, while workbook-only display
  columns remain local; and
- `Condition` is a managed inventory header even though other headers may be
  extended.

This is a greenfield cutover. There is no legacy `ROW` mapping, old-inventory
import, or compatibility requirement. Existing unmanaged business data is left
behind. Test inventory comes only from a newly generated warehouse or the
Admin `Seed Demo Inventory` tool.

### 5.2 Modeless role-form contract

The main Receiving, Production, and Shipping forms open modelessly so operators
can inspect and use their workbook while a role form remains open. Modal
confirmation, credential, or narrowly scoped selection dialogs remain permitted
when the workflow requires an explicit blocking decision.

Modeless operation must not make `ActiveWorkbook`, `ActiveSheet`, or
`Selection` authoritative. Each main role form captures its intended operator
workbook and typed workflow/session context when opened. Every action, refresh,
and render resolves through that captured context. If the workbook closes,
changes authority, or is no longer valid, the form disables write actions and
reports the condition instead of silently rebinding to another workbook.

Only one main form instance per role and captured operator context may own a
live workflow session. Repeated ribbon launches activate the existing valid
instance or create a new correctly bound instance; they must not create
duplicate hidden sessions or event subscriptions.

## 6. D13 execution protocol for every slice

Before implementation:

1. Name the contract being changed.
2. Name the focused test that protects it.
3. Run and record RED for the expected behavioral reason.
4. If no test exists, create it before editing runtime code.

After implementation:

1. Record focused GREEN.
2. Run the relevant role/package regression set.
3. Regenerate static maintenance evidence.
4. Compare bloat/dynamic-call metrics with the previous baseline.

A compile failure, missing fixture, broken harness, or unavailable workbook is not meaningful RED.

## 7. Revised slice sequence

### [x] Entry gate — normative `System_Key` decision

Before Slice 0 begins, amend the normative design specification and
`AGENTS.md` so they no longer define `ROW` or `(ITEM_CODE, Location)` as
canonical identity. The specification must define:

- [x] the `System_Key` generation, immutability, reference, and projection rules;
- [x] extensible managed/shared/local header behavior;
- [x] `Condition` as a managed inventory field;
- [x] the greenfield no-import boundary; and
- [x] the Admin/Create Warehouse seed acceptance contract.

D14 in the normative specification and the matching `AGENTS.md` invariants
satisfy this architectural entry gate. No runtime implementation of
`System_Key` begins before the amendment is reviewed together and the required
D13 RED tests are recorded.

### [x] Slice 0 — Tool contracts, schemas, and synthetic fixtures

Write failing tests for the scanner and runtime extractor before implementing either tool.

Deliver:

- [x] versioned JSON schemas;
- [x] synthetic VBA/Ribbon fixtures containing direct, dynamic, event, duplicate, and unreachable examples;
- [x] synthetic runtime workbook/add-in metadata fixtures;
- [x] synthetic table fixtures with `System_Key`, managed headers, unknown custom
  headers, and no `ROW` header;
- [x] redaction tests;
- [x] deterministic JSON-to-Markdown rendering tests; and
- [x] the dynamic-root registry format.

Gate:

- [x] fixtures demonstrate RED against the absent tools;
- [x] schemas distinguish static implementation data from runtime data;
- [x] scanner/runtime fixtures distinguish managed headers from preserved custom
  headers and flag `ROW` as a retired runtime contract; and
- [x] no operational workbook is required for tool unit tests.

### [x] Slice 1 — Static scanner MVP

Implement Tool A against Receiving, Production, Shipping, Core, Domain, and Admin sources.

Gate:

- [x] deterministic output on repeated runs;
- [x] all Ribbon, event, and literal macro roots are represented;
- [x] unresolved dynamic calls are reported, not guessed;
- [x] package composition matches the current build map;
- [x] JSON validates against its schema; and
- [x] Markdown is regenerated byte-for-byte from JSON.

### [x] Slice 2 — Runtime-state extractor MVP

Implement Tool B as a read-only diagnostic path.

Gate:

- [x] a before/after hash proves opened operational workbooks were not changed;
- [x] no workbook is saved, processed, repaired, refreshed, or closed;
- [x] secrets are redacted;
- [x] current legacy package layout is reported accurately; and
- [x] JSON/Markdown outputs agree.

### [x] Slice 3 — Baseline, root registry, and reviewed cleanup backlog

Generate the first complete baseline.

Classify candidates into:

- [x] remove;
- [x] move to tests;
- [x] split module;
- [x] replace duplicate;
- [x] replace same-project late binding;
- [x] retain as dynamic root;
- [x] remove or isolate abandoned legacy inventory/import paths; and
- [x] unresolved/manual investigation.

Gate:

- [x] no automatic deletions;
- [x] every HIGH-confidence deletion candidate has a reason and protecting test;
- [x] the backlog separately identifies Production, Receiving, Shipping, and shared-package work; and
- [x] module-growth ratchets are recorded.

### [x] Slice 4 — Greenfield warehouse generation and demo inventory seeding

Do not import, translate, or reconcile old operational inventory. Make the
supported clean-start tools authoritative:

- [x] Admin Generate Warehouse/Create Warehouse;
- [x] automatic bootstrap demo seed when selected; and
- [x] Admin `Seed Demo Inventory`.

Write meaningful RED tests before changing implementation. The packaged actions
must create a fresh warehouse whose Inventory Domain, snapshots, operator read
models, and fake inventory use `System_Key` and contain no `ROW` header or
dependency.

Gate:

- [x] every generated durable inventory entity has a nonblank globally unique
  `System_Key`;
- [x] repeated seed actions create new keys and do not collide with or overwrite
  earlier fake inventory;
- [x] generated Inventory Domain, snapshot, and operator tables contain the required
  managed headers and tolerate additional custom headers;
- [x] `Condition` is present and defaults seeded inventory to the allowed value
  `GOOD`;
- [x] processor application, snapshot publication, and operator refresh preserve
  each key;
- [x] no generated runtime table contains a `ROW` header;
- [x] no runtime generation/seed path calls legacy inventory import or migration
  logic; and
- [x] packaged Admin Generate Warehouse and `Seed Demo Inventory` tests are GREEN.

### [x] Slice 5 — Pre-refactor packaged behavior locks

This slice must precede service extraction.

Write and observe meaningful RED for:

- [x] Production two consecutive batches through the actual form actions;
- [x] Receiving Confirm Writes through the actual form action;
- [x] the main Receiving form exposing a selectable Purchasing stub tab that is
  clearly non-operational and performs no writes;
- [x] Shipping/Boxing through Shipments Sent using the actual form actions;
- [x] the main Receiving, Production, and Shipping forms opening modelessly while
  the captured operator workbook remains usable, without actions drifting to a
  subsequently activated workbook;
- [x] the Operations Shipping launcher opening one main Shipping form whose tabs
  include Box Builder and Box Maker, with no separate Box Builder or Box Maker
  buttons on the Operations ribbon; and
- [x] restart/reopen preservation at the highest-risk boundary currently failing.

These tests must use the same selection, Apply, Check In, Complete Run, Confirm Writes, To Shipments, Shipments Sent, refresh, and Next Batch handlers used by operators. Direct service calls are additional tests, not substitutes.

Gate:

- [x] RED failures correspond to known behavioral gaps;
- [x] current successful legacy behavior is characterized so refactoring does not erase it accidentally; and
- [x] the test harness can distinguish UI/controller failure from processor/domain failure.

### [x] Slice 6 — Shadow Operations package and collision harness

Create a non-deployed `invSys.Operations.xlam` shadow build early. Import the three role source sets without retiring or registering over the existing role add-ins.

Purpose:

- [x] discover duplicate module/form names;
- [x] discover public procedure collisions;
- [x] discover Ribbon callback collisions;
- [x] prove compile/reference/load order; and
- [x] establish the future package boundary before deep refactoring.

This is not the D12 cutover.

Gate:

- [x] shadow package compiles;
- [x] all three role forms initialize in isolation;
- [x] collision report is empty or explicitly resolved;
- [x] legacy add-ins remain the active operational package; and
- [x] no duplicate ribbon is registered for normal use.

### [x] Slice 7 — Production session and completion service

Write focused service/session RED tests in addition to the Slice 5 packaged RED test.

Implement:

- [x] typed Production session state;
- [x] canonical `System_Key` allocation identity;
- [x] structured completion result;
- [x] explicit consume/output event identities;
- [x] processor result verification;
- [x] snapshot/read-model refresh result;
- [x] ready-for-next-batch transition; and
- [x] failure/compensation semantics.

Gate:

- [x] no control or worksheet table is treated as the authoritative session object;
- [x] no canonical Domain workbook is mutated directly;
- [x] focused service tests are GREEN;
- [x] the packaged two-batch target advances toward GREEN without weakening assertions; and
- [x] session state survives the required restart boundary.

### [x] Slice 8 — Production controller, legacy retirement, and typed internal calls

Make `frmProduction` a thin controller/renderer over the Production session/service.

Remove or isolate:

- [x] implicit legacy recipe fallback when Designs is enabled;
- [x] every `ROW` header, lookup, display, hidden-list field, and authority path;
- [x] same-project `Application.Run`;
- [x] direct canonical workbook writes;
- [x] duplicate table/header helpers;
- [x] test fixtures embedded in runtime modules; and
- [x] diagnostics that no longer protect a live incident.

Gate:

- [x] Designs-enabled Production reads released Designs Domain recipes only;
- [x] the Production form opens modelessly and remains bound to its captured
  operator workbook/session while the operator uses other workbook surfaces;
- [x] direct typed calls are used inside the Operations project;
- [x] cross-XLAM bridges use declared primitive/JSON contracts;
- [x] two-batch packaged form-action test is GREEN;
- [x] scanner confirms the targeted legacy paths are absent; and
- [x] no runtime module exceeds its previous bloat baseline without an exception.

### [x] Slice 9 — Production layout standardization

Replace one-off coordinate arithmetic with the v4.11 Windows API plus anchor-based layout standard.

Do not run this in parallel with Slice 8 when both modify `frmProduction`.

Gate:

- [x] minimum/default/max geometry checks;
- [x] no overlap or out-of-bounds controls;
- [x] screenshots at required sizes;
- [x] minimize/maximize behavior;
- [x] no regression in selection or action handlers; and
- [x] layout metadata appears in the static manifest where practical.

### [x] Slice 10 — Receiving service and form stabilization

Define Receiving's own workflow state and typed posting service:

```text
staged -> validated -> submitted -> processor applied -> snapshot refreshed -> cleared/ready
```

Retire redundant worksheet mutation and legacy posting paths.

The main Receiving form also reserves a `Purchasing` tab as a visible,
selectable stub for future work. The stub must:

- [x] identify itself clearly as not yet operational;
- [x] contain no enabled purchasing write/post controls;
- [x] submit no events and mutate no workbook, Domain, or runtime state;
- [x] reuse the Receiving form shell rather than opening a separate form; and
- [x] create no Purchasing button, group, or launch surface on the Operations
  ribbon.

Adding the stub does not introduce a Purchasing service, capability, event type,
Domain contract, or implementation slice. Those require a future normative
design decision and their own test-first plan.

Gate:

- [x] D13 service RED/GREEN evidence;
- [x] packaged Confirm Writes form-action test GREEN;
- [x] the Receiving form opens modelessly and remains bound to its captured
  operator workbook/session;
- [x] packaged navigation test proves the Purchasing tab exists, is selectable,
  remains visibly non-operational, and performs no writes or event submission;
- [x] event identity/idempotency proof;
- [x] staging clears only after confirmed submission/application according to contract;
- [x] snapshot refresh is non-destructive; and
- [x] scanner confirms retired paths are gone.

### [x] Slice 11 — Shipping and Boxing service stabilization

Shipping is large enough to require its own slice. Preserve D11's A+B event loop while separating:

- [x] Box design/version selection;
- [x] local shipment staging and locks;
- [x] To Shipments transfer;
- [x] Shipments Sent event submission;
- [x] processor application;
- [x] NAS/read-model refresh; and
- [x] lock release/compensation.

The main Shipping form becomes the role shell, following the tabbed navigation
pattern established by the Production form. Shipping workflow surfaces that
were launched by separate Excel ribbon buttons move into tabs/pages of that
form. In particular:

- [x] Box Builder is a tab in the main Shipping form;
- [x] Box Maker is a tab in the main Shipping form;
- [x] switching tabs preserves the Shipping session and refreshes only the selected
  workflow surface according to its contract;
- [x] the same Shipping capability gates continue to protect each tab's write
  actions; and
- [x] the Operations ribbon launches the main Shipping form and does not expose
  separate Box Builder or Box Maker buttons.

This is a navigation and form-composition change. It does not merge Box Builder,
Box Maker, shipment staging, or Shipments Sent into one mutable workflow state.
Their controllers/services remain separated behind the tabbed role shell.

Gate:

- [x] D13 service RED/GREEN evidence;
- [x] packaged Shipping/Boxing form-action test GREEN;
- [x] the Shipping form opens modelessly and remains bound to its captured operator
  workbook/session while its tabs and workbook remain usable;
- [x] packaged navigation test proves the Shipping launcher opens the main Shipping
  form, the Box Builder and Box Maker tabs exist, and their existing operator
  actions remain reachable;
- [x] packaged RibbonX test proves the Operations ribbon contains no direct Box
  Builder or Box Maker buttons or callbacks;
- [x] `NAS Inv` is never mutated by local overlays;
- [x] `Projected Inv` derives only from D11 inputs;
- [x] Remove releases the exact active lock;
- [x] shipment replay is idempotent; and
- [x] restart does not resurrect completed staging or corrupt locks.

### [x] Slice 12 — Reviewed code-bloat cleanup gate

Run a dedicated cleanup only after role contracts are protected.

Priorities:

- [x] move test fixtures out of runtime modules;
- [x] remove HIGH-confidence unreachable compatibility code;
- [x] consolidate duplicate table/header/list helpers;
- [x] split oversized modules by responsibility;
- [x] replace string reports with structured internal results;
- [x] reduce unresolved late binding; and
- [x] delete obsolete diagnostics with replacement observability.

Gate:

- [x] every deletion maps to scanner evidence and a protecting test;
- [x] all role-focused tests and packaged smoke tests pass;
- [x] code-size and duplicate metrics improve;
- [x] no new dynamic-root exceptions are added merely to silence the scanner; and
- [x] generated reports explain retained legacy code.

### [x] Slice 13 — D12 final Operations cutover

Promote the shadow package to the deployed `invSys.Operations.xlam`.

Implement:

- [x] one Operations ribbon with independently gated Receiving, Production, and Shipping groups;
- [x] modeless main Receiving, Production, and Shipping form launchers with explicit
  operator-workbook/session binding and duplicate-instance prevention;
- [x] one Receiving launch surface whose main form includes the non-operational
  Purchasing stub tab, with no separate Purchasing ribbon button or group;
- [x] one Shipping launch surface on that ribbon; Box Builder and Box Maker are
  tabs in the main Shipping form rather than ribbon buttons;
- [x] shared connection/sign-in/status controls;
- [x] five-package build and manifest;
- [x] selective complete-Operations-project build;
- [x] standalone role-add-in unregister/removal;
- [x] coexistence detection;
- [x] duplicate callback/tab prevention; and
- [x] clean restart/load-order validation.

Gate:

- [x] exactly five published XLAMs;
- [x] operations-only account shows one invSys tab;
- [x] administrative setup shows at most Operations and Admin;
- [x] Core/Domain XLAMs remain headless;
- [x] legacy role XLAMs are absent;
- [x] each main role form is modeless, leaves its captured operator workbook usable,
  and cannot redirect an action to a different active workbook;
- [x] the main Receiving form exposes the Purchasing stub tab and the Operations
  ribbon exposes no Purchasing button or group;
- [x] the Operations ribbon has no Box Builder or Box Maker buttons and the main
  Shipping form exposes both workflows as tabs;
- [x] every role form compiles and initializes; and
- [x] failure diagnostics identify the role without silently disabling the other loaded surfaces.

### [x] Slice 14 — Full chain, restart, and reconciliation validation

Validate the actual operator chain against the consolidated package:

```text
Generate a fresh warehouse
-> Seed fake inventory through Admin
-> Receive inventory
-> processor apply
-> snapshot/read-model refresh
-> Production allocation and two batches
-> component consumption and output creation
-> Boxing/version selection
-> shipment staging and Shipments Sent
-> processor apply
-> final snapshot/read-model refresh
```

Required assertions:

- [x] every durable inventory entity carries a unique immutable `System_Key`;
- [x] `ROW` is absent from generated, canonical, snapshot, and operator runtime
  tables;
- [x] managed `Condition` and added custom headers survive the tested refresh and
  rebuild boundaries;
- [x] exact input/output quantities and locations;
- [x] event IDs, statuses, log rows, and idempotent replay;
- [x] no negative inventory;
- [x] batch number, Last, Total, output identity, and ready-for-next-batch state;
- [x] correct box/BOM version;
- [x] exact lock acquisition/release;
- [x] no local overlay corruption of NAS inventory;
- [x] restart/reopen at selected boundaries;
- [x] no visible canonical workbook pollution;
- [x] no duplicate tabs/add-ins/callback execution;
- [x] runtime extractor reports the expected five-package state; and
- [x] static scanner reports no reintroduced retired path.

## 8. Parallelism and ownership

Safe parallel work:

- Slice 1 static scanner and Slice 2 runtime extractor after Slice 0 schemas stabilize;
- Receiving and Shipping service design after shared result contracts stabilize, provided agents edit separate files;
- documentation renderer and JSON comparison tooling after schema freeze.

Unsafe parallel work:

- Production controller and Production layout changes in `frmProduction`;
- Operations RibbonX consolidation and legacy callback removal;
- multiple agents editing the build project map;
- cleanup deletion while another slice is moving the same procedures.

One agent or workstream owns each overlapping file set. Generated reports may be shared, but agents must not edit generated Markdown manually.

## 9. Completion definition

Operations stabilization is complete:

- [x] a fresh warehouse can be generated and seeded without importing legacy
  inventory;
- [x] the runtime uses `System_Key` and contains no `ROW` identity dependency;
- [x] managed and end-user-added headers survive their declared persistence
  boundaries;
- [x] D12 five-package deployment is active;
- [x] D13 RED/GREEN evidence exists for each changed service/form-action contract;
- [x] Receiving, Production, Boxing, and Shipping pass their packaged action tests;
- [x] the full chain passes with reconciliation and restart checks;
- [x] legacy direct-mutation/fallback paths identified for retirement are absent;
- [x] static and runtime JSON reports are reproducible and redacted;
- [x] code-bloat metrics are lower and ratcheted against regrowth; and
- [x] a new session can reconstruct current implementation/runtime state from generated evidence without relying on chat memory.
