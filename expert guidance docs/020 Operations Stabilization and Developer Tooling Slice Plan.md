# Operations Stabilization and Developer Tooling Slice Plan

**Status:** Revised implementation guidance
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

- stabilize Receiving, Production, Boxing, and Shipping as one coherent event-driven Operations system;
- remove legacy direct-mutation and fallback paths that conflict with the Domain architecture;
- reduce code bloat without deleting dynamic VBA entry points accidentally;
- extract current implementation and runtime state on demand as JSON and generated Markdown;
- consolidate the three role packages into `invSys.Operations.xlam` without creating another monolith; and
- prove the complete Receiving -> Production -> Boxing -> Shipping chain through the packaged operator actions.

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

Production identity should use `ITEM_CODE`/SKU plus Location for inventory allocations. Recipe/design identity, batch identity, event identity, and output identity remain separate fields. Legacy worksheet `ROW` may be retained only as migration/display metadata, never as the canonical session key.

### 5.1 Modeless role-form contract

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

### Slice 0 — Tool contracts, schemas, and synthetic fixtures

Write failing tests for the scanner and runtime extractor before implementing either tool.

Deliver:

- versioned JSON schemas;
- synthetic VBA/Ribbon fixtures containing direct, dynamic, event, duplicate, and unreachable examples;
- synthetic runtime workbook/add-in metadata fixtures;
- redaction tests;
- deterministic JSON-to-Markdown rendering tests; and
- the dynamic-root registry format.

Gate:

- fixtures demonstrate RED against the absent tools;
- schemas distinguish static implementation data from runtime data; and
- no operational workbook is required for tool unit tests.

### Slice 1 — Static scanner MVP

Implement Tool A against Receiving, Production, Shipping, Core, Domain, and Admin sources.

Gate:

- deterministic output on repeated runs;
- all Ribbon, event, and literal macro roots are represented;
- unresolved dynamic calls are reported, not guessed;
- package composition matches the current build map;
- JSON validates against its schema; and
- Markdown is regenerated byte-for-byte from JSON.

### Slice 2 — Runtime-state extractor MVP

Implement Tool B as a read-only diagnostic path.

Gate:

- a before/after hash proves opened operational workbooks were not changed;
- no workbook is saved, processed, repaired, refreshed, or closed;
- secrets are redacted;
- current legacy package layout is reported accurately; and
- JSON/Markdown outputs agree.

### Slice 3 — Baseline, root registry, and reviewed cleanup backlog

Generate the first complete baseline.

Classify candidates into:

- remove;
- move to tests;
- split module;
- replace duplicate;
- replace same-project late binding;
- retain as dynamic root;
- retain temporarily as legacy migration; and
- unresolved/manual investigation.

Gate:

- no automatic deletions;
- every HIGH-confidence deletion candidate has a reason and protecting test;
- the backlog separately identifies Production, Receiving, Shipping, and shared-package work; and
- module-growth ratchets are recorded.

### Slice 4 — Pre-refactor packaged behavior locks

This slice must precede service extraction.

Write and observe meaningful RED for:

- Production two consecutive batches through the actual form actions;
- Receiving Confirm Writes through the actual form action;
- the main Receiving form exposing a selectable Purchasing stub tab that is
  clearly non-operational and performs no writes;
- Shipping/Boxing through Shipments Sent using the actual form actions;
- the main Receiving, Production, and Shipping forms opening modelessly while
  the captured operator workbook remains usable, without actions drifting to a
  subsequently activated workbook;
- the Operations Shipping launcher opening one main Shipping form whose tabs
  include Box Builder and Box Maker, with no separate Box Builder or Box Maker
  buttons on the Operations ribbon; and
- restart/reopen preservation at the highest-risk boundary currently failing.

These tests must use the same selection, Apply, Check In, Complete Run, Confirm Writes, To Shipments, Shipments Sent, refresh, and Next Batch handlers used by operators. Direct service calls are additional tests, not substitutes.

Gate:

- RED failures correspond to known behavioral gaps;
- current successful legacy behavior is characterized so refactoring does not erase it accidentally; and
- the test harness can distinguish UI/controller failure from processor/domain failure.

### Slice 5 — Shadow Operations package and collision harness

Create a non-deployed `invSys.Operations.xlam` shadow build early. Import the three role source sets without retiring or registering over the existing role add-ins.

Purpose:

- discover duplicate module/form names;
- discover public procedure collisions;
- discover Ribbon callback collisions;
- prove compile/reference/load order; and
- establish the future package boundary before deep refactoring.

This is not the D12 cutover.

Gate:

- shadow package compiles;
- all three role forms initialize in isolation;
- collision report is empty or explicitly resolved;
- legacy add-ins remain the active operational package; and
- no duplicate ribbon is registered for normal use.

### Slice 6 — Production session and completion service

Write focused service/session RED tests in addition to the Slice 4 packaged RED test.

Implement:

- typed Production session state;
- canonical allocation identity;
- structured completion result;
- explicit consume/output event identities;
- processor result verification;
- snapshot/read-model refresh result;
- ready-for-next-batch transition; and
- failure/compensation semantics.

Gate:

- no control or worksheet table is treated as the authoritative session object;
- no canonical Domain workbook is mutated directly;
- focused service tests are GREEN;
- the packaged two-batch target advances toward GREEN without weakening assertions; and
- session state survives the required restart boundary.

### Slice 7 — Production controller, legacy retirement, and typed internal calls

Make `frmProduction` a thin controller/renderer over the Production session/service.

Remove or isolate:

- implicit legacy recipe fallback when Designs is enabled;
- legacy `ROW` authority;
- same-project `Application.Run`;
- direct canonical workbook writes;
- duplicate table/header helpers;
- test fixtures embedded in runtime modules; and
- diagnostics that no longer protect a live incident.

Gate:

- Designs-enabled Production reads released Designs Domain recipes only;
- the Production form opens modelessly and remains bound to its captured
  operator workbook/session while the operator uses other workbook surfaces;
- direct typed calls are used inside the Operations project;
- cross-XLAM bridges use declared primitive/JSON contracts;
- two-batch packaged form-action test is GREEN;
- scanner confirms the targeted legacy paths are absent; and
- no runtime module exceeds its previous bloat baseline without an exception.

### Slice 8 — Production layout standardization

Replace one-off coordinate arithmetic with the v4.11 Windows API plus anchor-based layout standard.

Do not run this in parallel with Slice 7 when both modify `frmProduction`.

Gate:

- minimum/default/max geometry checks;
- no overlap or out-of-bounds controls;
- screenshots at required sizes;
- minimize/maximize behavior;
- no regression in selection or action handlers; and
- layout metadata appears in the static manifest where practical.

### Slice 9 — Receiving service and form stabilization

Define Receiving's own workflow state and typed posting service:

```text
staged -> validated -> submitted -> processor applied -> snapshot refreshed -> cleared/ready
```

Retire redundant worksheet mutation and legacy posting paths.

The main Receiving form also reserves a `Purchasing` tab as a visible,
selectable stub for future work. The stub must:

- identify itself clearly as not yet operational;
- contain no enabled purchasing write/post controls;
- submit no events and mutate no workbook, Domain, or runtime state;
- reuse the Receiving form shell rather than opening a separate form; and
- create no Purchasing button, group, or launch surface on the Operations
  ribbon.

Adding the stub does not introduce a Purchasing service, capability, event type,
Domain contract, or implementation slice. Those require a future normative
design decision and their own test-first plan.

Gate:

- D13 service RED/GREEN evidence;
- packaged Confirm Writes form-action test GREEN;
- the Receiving form opens modelessly and remains bound to its captured
  operator workbook/session;
- packaged navigation test proves the Purchasing tab exists, is selectable,
  remains visibly non-operational, and performs no writes or event submission;
- event identity/idempotency proof;
- staging clears only after confirmed submission/application according to contract;
- snapshot refresh is non-destructive; and
- scanner confirms retired paths are gone.

### Slice 10 — Shipping and Boxing service stabilization

Shipping is large enough to require its own slice. Preserve D11's A+B event loop while separating:

- Box design/version selection;
- local shipment staging and locks;
- To Shipments transfer;
- Shipments Sent event submission;
- processor application;
- NAS/read-model refresh; and
- lock release/compensation.

The main Shipping form becomes the role shell, following the tabbed navigation
pattern established by the Production form. Shipping workflow surfaces that
were launched by separate Excel ribbon buttons move into tabs/pages of that
form. In particular:

- Box Builder is a tab in the main Shipping form;
- Box Maker is a tab in the main Shipping form;
- switching tabs preserves the Shipping session and refreshes only the selected
  workflow surface according to its contract;
- the same Shipping capability gates continue to protect each tab's write
  actions; and
- the Operations ribbon launches the main Shipping form and does not expose
  separate Box Builder or Box Maker buttons.

This is a navigation and form-composition change. It does not merge Box Builder,
Box Maker, shipment staging, or Shipments Sent into one mutable workflow state.
Their controllers/services remain separated behind the tabbed role shell.

Gate:

- D13 service RED/GREEN evidence;
- packaged Shipping/Boxing form-action test GREEN;
- the Shipping form opens modelessly and remains bound to its captured operator
  workbook/session while its tabs and workbook remain usable;
- packaged navigation test proves the Shipping launcher opens the main Shipping
  form, the Box Builder and Box Maker tabs exist, and their existing operator
  actions remain reachable;
- packaged RibbonX test proves the Operations ribbon contains no direct Box
  Builder or Box Maker buttons or callbacks;
- `NAS Inv` is never mutated by local overlays;
- `Projected Inv` derives only from D11 inputs;
- Remove releases the exact active lock;
- shipment replay is idempotent; and
- restart does not resurrect completed staging or corrupt locks.

### Slice 11 — Reviewed code-bloat cleanup gate

Run a dedicated cleanup only after role contracts are protected.

Priorities:

1. move test fixtures out of runtime modules;
2. remove HIGH-confidence unreachable compatibility code;
3. consolidate duplicate table/header/list helpers;
4. split oversized modules by responsibility;
5. replace string reports with structured internal results;
6. reduce unresolved late binding; and
7. delete obsolete diagnostics with replacement observability.

Gate:

- every deletion maps to scanner evidence and a protecting test;
- all role-focused tests and packaged smoke tests pass;
- code-size and duplicate metrics improve;
- no new dynamic-root exceptions are added merely to silence the scanner; and
- generated reports explain retained legacy code.

### Slice 12 — D12 final Operations cutover

Promote the shadow package to the deployed `invSys.Operations.xlam`.

Implement:

- one Operations ribbon with independently gated Receiving, Production, and Shipping groups;
- modeless main Receiving, Production, and Shipping form launchers with explicit
  operator-workbook/session binding and duplicate-instance prevention;
- one Receiving launch surface whose main form includes the non-operational
  Purchasing stub tab, with no separate Purchasing ribbon button or group;
- one Shipping launch surface on that ribbon; Box Builder and Box Maker are
  tabs in the main Shipping form rather than ribbon buttons;
- shared connection/sign-in/status controls;
- five-package build and manifest;
- selective complete-Operations-project build;
- standalone role-add-in unregister/removal;
- coexistence detection;
- duplicate callback/tab prevention; and
- clean restart/load-order validation.

Gate:

- exactly five published XLAMs;
- operations-only account shows one invSys tab;
- administrative setup shows at most Operations and Admin;
- Core/Domain XLAMs remain headless;
- legacy role XLAMs are absent;
- each main role form is modeless, leaves its captured operator workbook usable,
  and cannot redirect an action to a different active workbook;
- the main Receiving form exposes the Purchasing stub tab and the Operations
  ribbon exposes no Purchasing button or group;
- the Operations ribbon has no Box Builder or Box Maker buttons and the main
  Shipping form exposes both workflows as tabs;
- every role form compiles and initializes; and
- failure diagnostics identify the role without silently disabling the other loaded surfaces.

### Slice 13 — Full chain, restart, and reconciliation validation

Validate the actual operator chain against the consolidated package:

```text
Receive inventory
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

- exact input/output quantities and locations;
- event IDs, statuses, log rows, and idempotent replay;
- no negative inventory;
- batch number, Last, Total, output identity, and ready-for-next-batch state;
- correct box/BOM version;
- exact lock acquisition/release;
- no local overlay corruption of NAS inventory;
- restart/reopen at selected boundaries;
- no visible canonical workbook pollution;
- no duplicate tabs/add-ins/callback execution;
- runtime extractor reports the expected five-package state; and
- static scanner reports no reintroduced retired path.

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

Operations stabilization is complete only when:

- D12 five-package deployment is active;
- D13 RED/GREEN evidence exists for each changed service/form-action contract;
- Receiving, Production, Boxing, and Shipping pass their packaged action tests;
- the full chain passes with reconciliation and restart checks;
- legacy direct-mutation/fallback paths identified for retirement are absent;
- static and runtime JSON reports are reproducible and redacted;
- code-bloat metrics are lower and ratcheted against regrowth; and
- a new session can reconstruct current implementation/runtime state from generated evidence without relying on chat memory.
