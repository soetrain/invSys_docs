# Operations Launchers Open and Seed Demo Inventory Blocks Control UAT

**Created:** 2026-08-02

## Goal and release outcome

Achieve Release 1 user acceptance under `invSys-Design-v4.11.md` by preserving
the now-working packaged Operations launchers and diagnosing the first failed
downstream control, Admin **Seed Demo Inventory**, before continuing batched
control-by-control acceptance.

This advances the v4.11 Phase 6 saved-workbook/NAS acceptance and the D13
requirement that packaged operator actions receive meaningful behavioral
RED/GREEN coverage.

## Current verified state

- Code repository: branch `main`, commit
  `81ce43667bd2326d9a6b9cafd59413efd0656e3a`, tracking `origin/main` on
  `soetrain/invSys_fork`; working tree clean before this handoff. Last verified:
  2026-08-02.
- Documentation repository: branch `main`, commit
  `4885033b58b699d313133905147ae6f1f7c9dcfa`, tracking `origin/main` on
  `soetrain/invSys_docs`; working tree clean before this handoff. Last verified:
  2026-08-02.
- Active work remains Plan 022 Slice 4 and its batched user checkpoint. Its
  automated isolated, packaged, restart, and dedicated NAS gates are GREEN.
- User acceptance on 2026-08-02 confirmed that the Receiving, Production, and
  Shipping forms all opened and appeared usable. The complete restart,
  captured-workbook, and repeat-launch checklist has not yet been returned, so
  Plan 022 is not complete.
- The next action, Admin **Seed Demo Inventory**, caused Excel to flash for a
  period and then display `Application-defined or object-defined error`.
- Core and Domain XLAMs did not initially appear loaded in Excel. The user
  manually added them and retried; **Seed Demo Inventory** still failed.
- Treat Excel as open. Close it before rebuilding or replacing XLAMs.

## Decisions and constraints

- Treat the seed failure as a newly discovered Release 1 acceptance blocker,
  not as a reason to reopen the successful launcher implementation without
  evidence.
- Assess and reproduce the failure through the packaged public ribbon callback
  `modAdmin.Seed_DemoInventory` before choosing a targeted fix or refactor.
- The missing Core/Domain load state is relevant evidence but is not a
  confirmed root cause because manually loading those add-ins did not repair
  the action.
- Continue control testing in bounded batches after this failure is GREEN; do
  not assume that opening each form proves its form actions.
- Use only the dedicated generated NAS test warehouse for write-capable tests.
- Preserve D12 packaging, Domain authority, `System_Key`, extensible headers,
  and station-local operator-workbook boundaries.

## Evidence and traceability

- Final Plan 022 automated evidence in
  `tests/integration/plan022_results.md` records launcher contracts 24/24,
  packaged validation 54/54, the full isolated launcher matrix 17/17, Release
  1 chain 30/30, and dedicated NAS validation 12/12.
- The user-visible launcher failures recorded in handoff 005 no longer
  reproduce in the current user session: all three forms opened.
- New user-visible failure:
  symptom -> Excel flashes, then displays `Application-defined or
  object-defined error` after **Seed Demo Inventory**;
  known/suspected boundary -> the ribbon maps to
  `modAdmin.Seed_DemoInventory`, which resolves interactive context and calls
  `modAdminInventorySeed.SeedDemoInventoryForWarehouse`; exact failing stage,
  error number, source, and loaded dependency state remain unresolved;
  governing requirement -> v4.11 requires Admin **Seed Demo Inventory** and
  D13 RED/GREEN coverage;
  required test -> packaged invocation of the same public ribbon callback with
  stage/error capture, supplemented by the automation/service seed test.
- Existing service-level evidence is not sufficient for this defect:
  `modAdminConsole.SeedDemoInventoryForAutomation` passed in the Release 1 and
  dedicated NAS automated chains, while the interactive public callback failed.

## Do Not Repeat

- Do not treat the passing automation helper as proof that the visible Admin
  ribbon action works.
- Do not infer that manually loading Core and Domain repaired or identified the
  defect; the retry still failed.
- Do not broaden into a control refactor before the packaged interactive RED
  identifies the failing stage.
- Do not rebuild or replace add-ins while Excel has them loaded.
- Do not run write-capable control tests against a non-test warehouse.

## Assumptions to Re-verify

- Re-verify the exact loaded add-in paths, hashes, load order, and registry
  startup entries in the user's failing Excel session.
- Re-verify that the selected target is the dedicated generated NAS test
  warehouse and that the signed-in user has `ADMIN_MAINT`.
- Re-verify whether the seed-selection form appeared before Excel began
  flashing and which workbook was active.
- Re-verify whether the failed action queued or applied any event before
  retrying; do not assume the call was atomic.
- Re-verify the full Plan 022 repeat-launch and restart checklist after the
  seed blocker is resolved.

## Open questions and blockers

- The exact failing stage, `Err.Number`, `Err.Source`, and mutated/unchanged
  runtime artifacts are unresolved.
- It is unresolved why Core and Domain were not initially shown as loaded, and
  whether the Admin dependency bootstrap used the approved deployed paths.
- Release 1 control acceptance cannot continue reliably until the seed action
  is reproduced and corrected.

## Immediate next action

With the current Excel session left read-only, capture redacted loaded-package
and target state plus before hashes, then create and run a focused packaged RED
through `modAdmin.Seed_DemoInventory` that reports the failing context, form,
queue, processor, or cleanup stage without depending only on the automation
helper.

## Critical references

- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
- `tests/integration/plan022_results.md`
- `src/Admin/Modules/modAdmin.bas`
- `src/Admin/Modules/modAdminInventorySeed.bas`
- `src/Admin/Modules/modAdminConsole.bas`
- `src/Core/Modules/modRoleEventWriter.bas`
- `src/Core/Modules/modProcessor.bas`
- `tests/integration/test_CreateWarehouse.bas`
- `tools/validate_release1_full_chain.ps1`
- `tools/validate_plan022_nas_runtime.ps1`
- Ribbon control `btnAdminSeedInventory`
- Public callback `modAdmin.Seed_DemoInventory`
- Automation helper `modAdminConsole.SeedDemoInventoryForAutomation`

