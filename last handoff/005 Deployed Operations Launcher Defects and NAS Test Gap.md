# Deployed Operations Launcher Defects and NAS Test Gap

**Created:** 2026-07-28

## Goal and release outcome

Restore the three deployed Operations ribbon launchers under an actual
NAS-connected operator session by adding Receiving station-local workbook
self-provisioning, repairing the Production and Shipping type mismatches, and
proving the result first in isolation and then against a dedicated NAS test
runtime.

This advances the v4.11 operator deployment model, station-local operator
context, D12 packaged launcher contract, D13 test-first rule, and Phase 6
saved-workbook/NAS acceptance.

## Current verified state

### Code repository

- Branch: `codex/fix-tester-station-nas-setup`
- HEAD: `b6ee0e91bd8c5e63c06d616fa28215c8e6b82715`
- Remote: `origin/codex/fix-tester-station-nas-setup` on
  `soetrain/invSys_fork`
- `origin/main` was at the same commit when last verified.
- Working tree was clean before this documentation-only continuation.
- Last verified: 2026-07-28.

### Documentation repository

- Branch: `agent/update-v4-11-operations-packaging`
- Current plan commit: `8daacdc` — `docs: add deployed launcher stabilization
  plan`
- Remote tracking branch:
  `fork/agent/update-v4-11-operations-packaging` on
  `justinwj/invSys_docs`
- Existing draft PR:
  `https://github.com/soetrain/invSys_docs/pull/1`
- `expert guidance docs/CURRENT.md` points to
  `022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`.
- Active plan state: Entry gate pending; Slices 0-4 pending.
- Last verified: 2026-07-28.

### Excel and runtime

- Excel was open for the user's reported connected/signed-in operator test.
  Treat it as open until re-verified.
- Excel must be closed before rebuilding, replacing, or loading a new XLAM
  package set in an automated Excel instance.
- The actually loaded XLAM paths/hashes, selected redacted UNC target, active
  workbook classification, warehouse/station context, and exact
  Production/Shipping `Err.Number`/`Err.Source` have not yet been captured.

## Decisions and constraints

- Assess and reproduce the defects before choosing an implementation strategy.
  Refactoring is allowed when the evidence and protecting tests show it is the
  right correction.
- The Receiving control must create or open the user's station-local Receiving
  operator workbook when no eligible workbook is open, then bind and show the
  modeless Receiving form.
- Operator workbooks remain station-local projections/staging surfaces. They
  are not canonical NAS workbooks.
- Production and Shipping must not return `Type mismatch`; the exact mismatch
  must be isolated by packaged-callback stage evidence before implementation.
- Do not infer that Production and Shipping require the same provisioning
  implementation as Receiving. Preserve their role-specific contracts.
- Do not select config, auth, inventory, snapshot, inbox, outbox, Domain XLAM,
  Operations XLAM, or an arbitrary active workbook as operator authority.
- D13 requires the actual callbacks
  `modTS_Received.ShowReceivingForm`,
  `mProduction.BtnOpenProductionForm`, and
  `modTS_Shipments.BtnOpenShipmentsForm` to produce meaningful behavioral RED
  before runtime edits.
- Automated NAS acceptance must use a dedicated generated test warehouse, not
  a warehouse containing non-test operational inventory.
- Runtime capture must be read-only and redacted.

## Evidence and traceability

User-visible acceptance evidence from 2026-07-28:

- Receiving after Connect Server and invSys sign-in:
  `Open a Receiving operator workbook before using the Receiving form.`
- Production:
  `Production form failed: Type mismatch`
- Shipping:
  `SHIPMENTS failed: Type mismatch`

Source observations, not confirmed root causes:

- `src/Receiving/Modules/modTS_Received.bas`:
  `ShowReceivingForm` resolves only the active non-add-in workbook and exits
  with the observed message when none is eligible.
- `src/Production/Modules/mProduction.bas`:
  `BtnOpenProductionForm` calls `ShowProductionForm ActiveWorkbook`;
  resolution contains active/bound/`ThisWorkbook` fallbacks and reports only
  the generic error description.
- `src/Shipping/Modules/modTS_Shipments.bas`:
  `BtnOpenShipmentsForm` resolves a workbook, begins quiet UI, repairs surfaces,
  initializes the form, and reports only the generic error description.
  The exact failing stage is unknown.

Verified test-scope gap:

- `tools/validate_release1_full_chain.ps1` loads the five XLAMs from
  `deploy/current`.
- It creates its runtime under
  `%TEMP%\invsys-release1-chain-<guid>`.
- Therefore plan 020 Slice 14's 30/30 result is packaged/isolation evidence,
  not actual UNC/NAS deployment evidence.
- No Excel/VBA tests were run and no operational workbook was changed while
  creating plan 022 and this handoff.

Symptom-to-test mapping:

| Symptom | Known/suspected cause | Requirement | Required protecting test |
|---|---|---|---|
| Receiving refuses to open without a workbook | Confirmed launcher exits instead of provisioning | v4.11 station-local operator context plus the user's binding decision | Packaged Receiving callback with no eligible workbook creates/reuses one saved station-local workbook and opens the captured modeless form |
| Production reports `Type mismatch` | Unresolved; broad workbook fallback or form/surface boundary is suspected | D12 modeless captured-workbook launcher | Packaged Production callback reports the exact failing stage in RED, then opens against an eligible saved workbook without mismatch |
| Shipping reports `Type mismatch` | Unresolved; workbook resolution/quiet UI/surface/form boundary is suspected | D11/D12 Shipping form boundary | Packaged Shipping callback reports the exact failing stage in RED, then opens against an eligible saved workbook without mismatch |
| Isolated tests passed but deployed run failed | Confirmed environment coverage gap | Phase 6 actual deployment acceptance | Hash-identified package set plus dedicated real NAS test-runtime checkpoint after isolated GREEN |

## Do Not Repeat

- Do not claim actual NAS validation from a test that uses a `%TEMP%` runtime.
- Do not start coding or commit to a local-patch versus refactor strategy before
  focused RED identifies the defects and affected boundaries.
- Do not treat `Type mismatch` alone as sufficient diagnostic evidence; capture
  error number, source, and initialization stage.
- Do not repair the symptom by accepting any active non-add-in workbook.
- Do not create station-local operator workbooks inside the canonical NAS
  runtime.
- Do not run write tests against a non-test operational warehouse.
- Do not build or replace XLAMs while Excel has them loaded.

## Assumptions to Re-verify

- Re-verify the code and documentation branch hashes above.
- Re-verify whether Excel remains open.
- Re-verify the loaded XLAM full paths/hashes and their relationship to
  `deploy/current`.
- Re-verify the selected warehouse target and station ID using redacted,
  read-only diagnostics.
- Re-verify whether saved Production or Shipping operator workbooks already
  exist and how the current station bootstrap records them.
- Re-verify that the NAS target used for acceptance is a dedicated test
  warehouse before any write-capable test.

## Open questions and blockers

- The exact Production and Shipping mismatch stages and error sources are
  unresolved; read-only diagnostics and focused packaged RED are required.
- Excel closure is not required for documentation or read-only source/runtime
  inspection, but it will be required before the first XLAM rebuild/deployment.
- No normative ambiguity or credential blocker is currently known.

## Immediate next action

Run the plan 022 Entry Gate read-only capture of loaded package hashes, selected
redacted target, and workbook classifications, then create the three packaged
launcher RED cases before changing runtime code.

## Critical references

- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`
- `src/Receiving/Modules/modTS_Received.bas`
- `src/Receiving/Forms/frmReceiving.frm`
- `src/Production/Modules/mProduction.bas`
- `src/Production/Forms/frmProduction.frm`
- `src/Shipping/Modules/modTS_Shipments.bas`
- `src/Shipping/Forms/frmShipmentsTally.frm`
- `src/Core/Modules/modRoleWorkbookSurfaces.bas`
- `src/Core/Modules/modWarehouseBootstrap.bas`
- `tools/validate_release1_full_chain.ps1`
- `tools/export-invsys-runtime-state.ps1`
- `tests/tooling/Test-Slice5BehaviorLocks.ps1`
- `tests/tooling/Test-Slice13OperationsCutover.ps1`
- Operations callbacks:
  `modTS_Received.ShowReceivingForm`,
  `mProduction.BtnOpenProductionForm`,
  `modTS_Shipments.BtnOpenShipmentsForm`
