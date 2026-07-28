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
| Production | `Production form failed: Type mismatch` | Open the Production form against an eligible captured operator workbook, or return a precise actionable role-workbook message without a VBA type mismatch. |
| Shipping | `SHIPMENTS failed: Type mismatch` | Open the Shipping form against an eligible captured operator workbook, or return a precise actionable role-workbook message without a VBA type mismatch. |

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

- [ ] complete the read-only deployed-runtime entry gate;
- [ ] reproduce all three failures through the packaged callbacks;
- [ ] capture the exact failing stage, error number, sanitized source, and
  description;
- [ ] compare the loaded package hashes with the intended deployed package;
- [ ] inspect the complete workbook-resolution and role-workbook provisioning
  paths; and
- [ ] identify whether the defects are isolated mistakes or symptoms of a
  broader launcher/context design problem.

Refactoring is allowed when the assessment and protecting RED demonstrate that
it is the safer or more maintainable correction. A justified refactor may
consolidate duplicated launcher/context logic, split responsibilities, or
change the affected form/controller boundary. It must:

- [ ] preserve the normative D12 package, Domain authority, event,
  `System_Key`, extensible-header, and captured-workbook contracts;
- [ ] state why a local patch is insufficient;
- [ ] define the bounded files, procedures, and contracts being changed before
  implementation;
- [ ] add focused tests for every contract moved or generalized;
- [ ] remain reversible and preserve plan 020 regression evidence;
- [ ] never accept an arbitrary active workbook, canonical runtime workbook, or
  XLAM as operator authority;
- [ ] never use a non-test operational NAS warehouse for write-capable tests;
  and
- [ ] never build or deploy XLAMs while Excel has the relevant add-ins or
  workbooks open.

## 5. D13 execution sequence

### Entry gate — identify the deployed runtime without mutation

Use the read-only runtime extractor and targeted source inspection to record:

- [ ] the full paths and hashes of the five loaded XLAMs;
- [ ] whether those hashes match the approved `deploy/current` manifest;
- [ ] the selected warehouse ID, station ID, target type, and redacted UNC/NAS
  runtime root;
- [ ] the active/open workbook classifications, without row-level operational
  values;
- [ ] whether an eligible station-local role workbook already exists for each
  launcher;
- [ ] the exact public callbacks:
  `modTS_Received.ShowReceivingForm`,
  `mProduction.BtnOpenProductionForm`, and
  `modTS_Shipments.BtnOpenShipmentsForm`; and
- [ ] before/after hashes proving the read-only capture changed no inspected
  workbook.

This gate must not save, refresh, repair, process, close, or create any
operational workbook.

### Slice 0 — packaged launcher RED and stage diagnostics

Create a focused packaged test that invokes the same three public callbacks as
the Operations ribbon under these states:

- [ ] connected and signed in, with no eligible role operator workbook open;
- [ ] a config/auth/snapshot/canonical workbook active;
- [ ] a saved eligible role workbook active;
- [ ] an unrelated ordinary workbook active; and
- [ ] a captured role workbook closed after form creation.

Record meaningful RED for:

- [ ] Receiving refusing to provision/open its station-local workbook;
- [ ] Production returning the observed type mismatch; and
- [ ] Shipping returning the observed type mismatch.

Add only enough safe diagnostic structure to report the failing launcher stage,
`Err.Number`, sanitized `Err.Source`, and `Err.Description`. A compile failure,
missing workbook, unavailable NAS, or broken harness is not RED.

Gate:

- [ ] all three failures reproduce through the packaged callbacks;
- [ ] the harness distinguishes resolution, provisioning, surface repair, form
  initialization, render, and modeless-show stages;
- [ ] no canonical or unrelated workbook is selected as operator authority; and
- [ ] the RED report contains no secrets or row-level operational data.

### Slice 1 — Receiving station-local workbook self-provisioning

Implement a narrow Receiving launcher path that:

- [ ] first reuses an already open eligible Receiving operator workbook;
- [ ] otherwise opens the existing station-local Receiving operator workbook
  resolved from the current warehouse/station context;
- [ ] otherwise creates and saves that workbook through a declared Core-owned
  primitive boundary;
- [ ] creates the managed Receiving and inventory read-model surfaces with no
  `ROW` header;
- [ ] refreshes from the selected warehouse snapshot without making the
  operator workbook authoritative;
- [ ] binds the modeless Receiving form to the created/opened workbook;
- [ ] activates or presents the operator workbook without changing the selected
  NAS target or signed-in invSys user;
- [ ] never saves or mutates canonical config, auth, inventory, snapshot,
  inbox, outbox, or Domain workbooks; and
- [ ] returns a precise recovery message if creation/opening cannot complete.

The path must be idempotent: repeated Receiving-control clicks reuse the same
eligible saved workbook and valid form instance rather than creating duplicate
files, forms, or event subscriptions.

Gate:

- [ ] meaningful packaged RED/GREEN is recorded;
- [ ] the actual Receiving ribbon callback creates/opens the workbook;
- [ ] the workbook is station-local and saved as a supported macro-enabled
  format;
- [ ] the form remains bound to its captured workbook when another workbook is
  activated;
- [ ] custom local columns survive reopening and refresh;
- [ ] no canonical NAS workbook hash changes merely from opening the form; and
- [ ] plan 020 Receiving, packaged XLAM, RibbonX, static-tooling, and
  maintenance regressions remain GREEN.

### Slice 2 — Production launcher type-mismatch repair

Use Slice 0 stage evidence to repair only the failing Production boundary.

Required behavior:

- [ ] `mProduction.BtnOpenProductionForm` validates the target before surface
  repair or form initialization;
- [ ] the callback never falls back to `ThisWorkbook` or another XLAM as
  operator authority;
- [ ] config, auth, inventory, snapshot, inbox, outbox, and Domain workbooks
  are rejected as operator workbooks;
- [ ] surface repair receives the declared primitive/workbook type at the
  correct project boundary;
- [ ] `frmProduction` initializes and opens modelessly against one captured
  eligible workbook; and
- [ ] a missing eligible workbook produces a role-specific actionable result,
  not `Type mismatch`.

Gate:

- [ ] focused packaged RED/GREEN identifies and removes the exact mismatch;
- [ ] saved-workbook initialize, captured-workbook, two-batch, restart, and
  layout regressions remain GREEN;
- [ ] no new same-project `Application.Run` or broad error suppression is
  introduced; and
- [ ] static bloat/dynamic-call metrics do not regress without a reviewed
  exception.

### Slice 3 — Shipping launcher type-mismatch repair

Use Slice 0 stage evidence to repair only the failing Shipping boundary.

Required behavior:

- [ ] `modTS_Shipments.BtnOpenShipmentsForm` validates an eligible Shipping
  operator workbook before quiet UI, surface repair, clipboard handling, form
  initialization, or autosync registration;
- [ ] config, auth, inventory, snapshot, inbox, outbox, and XLAM workbooks are
  rejected as operator authority;
- [ ] `Nothing` is never passed to a typed workbook operation;
- [ ] `frmShipmentsTally` initializes and opens modelessly against one captured
  eligible workbook; and
- [ ] a missing eligible workbook produces a role-specific actionable result,
  not `Type mismatch`.

Gate:

- [ ] focused packaged RED/GREEN identifies and removes the exact mismatch;
- [ ] Boxing tabs, Shipments staging, exact lock release, Shipments Sent,
  replay, restart, and overlay regressions remain GREEN;
- [ ] no canonical NAS workbook is polluted or selected as the operator
  workbook; and
- [ ] static bloat/dynamic-call metrics do not regress without a reviewed
  exception.

### Slice 4 — deployed-package and dedicated NAS test-runtime acceptance

After isolated RED/GREEN and regressions pass:

- [ ] rebuild the five-package set with Excel closed;
- [ ] record manifest paths and hashes;
- [ ] install/load the approved package set through the real account-scoped
  deployment path;
- [ ] use a dedicated generated test warehouse on the actual NAS/UNC path;
- [ ] keep station-local operator workbooks local to the user profile;
- [ ] connect, sign in, and invoke Receiving, Production, and Shipping through
  the visible Operations ribbon controls;
- [ ] restart Excel and repeat the launcher checks; and
- [ ] compare before/after hashes so only the dedicated test runtime and
  intended station-local operator workbooks changed.

Do not point this acceptance run at a warehouse containing non-test operational
inventory.

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
6. Click **Production** and confirm the form opens without a type mismatch.
7. Click **Shipping** and confirm the form opens without a type mismatch.
8. Activate a different ordinary workbook and confirm each open role form
   remains bound to its original operator workbook.
9. Close and reopen Excel, reconnect/sign in, and repeat steps 4-7.

Expected results:

- Receiving creates or reuses exactly one station-local saved operator
  workbook and opens its modeless form.
- Production and Shipping open modelessly or return a precise missing-role-
  workbook recovery instruction; neither returns `Type mismatch`.
- No config, auth, inventory, snapshot, inbox, outbox, or XLAM workbook becomes
  visible as an operator workbook.
- Repeated clicks and restart do not create duplicate forms, workbooks, tabs,
  add-ins, or callback execution.

Evidence to return:

- the full text of any dialog;
- whether each form opened;
- the operator workbook filename only, not its full user path;
- whether the second launch reused the same workbook/form; and
- screenshots of the Operations ribbon and each successful form, with
  sensitive warehouse/user data redacted.

## 7. Completion definition

This corrective plan is complete only when:

- [ ] the read-only deployed runtime evidence identifies the actually loaded
  package hashes and selected NAS test root;
- [ ] all three packaged launcher REDs are recorded for the observed behavioral
  reasons;
- [ ] Receiving self-provisions or opens its station-local workbook;
- [ ] Production and Shipping have no launcher type mismatch;
- [ ] plan 020 role, packaging, restart, static, and bloat regressions remain
  GREEN;
- [ ] the dedicated NAS test-runtime checkpoint passes after clean Excel
  restart; and
- [ ] the user returns the batched acceptance evidence.
