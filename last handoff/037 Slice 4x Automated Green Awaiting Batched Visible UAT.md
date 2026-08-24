# 037 Slice 4x Automated Green Awaiting Batched Visible UAT

## Goal and release outcome

Complete invSys Release 1 acceptance under Architecture v4.11 by obtaining the
remaining batched visible Plan 022 checkpoint after reusable Production and all
automated saved-workbook/NAS gates became GREEN.

## Current verified state

- Last verified: 2026-08-23.
- Code: `main` at `fd1b497`; docs: `main` at `079090c`; both pushed to `origin`.
- Active slice: Plan 022 Slice 4x, reusable Production Processes and Recipe
  graphs.
- Packaged Process/Recipe handlers, exact-key multi-output runs, clean Excel
  restart, and dedicated-NAS launcher readiness are complete.
- Excel is closed. Startup registration points to `deploy/current` Operations
  and Admin leaf XLAMs; Core and Domain packages remain headless/bridge-loaded.
- Unrelated user changes remain uncommitted and must be preserved: code
  `AGENTS.md` and docs `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`.

## Decisions and constraints

- Architecture v4.11 D15 is normative: Process Designer and Recipe Designer
  replace the operator-visible Recipe Builder; Run Tree remains experimental.
- Operator acceptance must use the public packaged launchers and visible form
  handlers against dedicated test warehouse `WHT7025AE`.
- Do not claim Release 1 acceptance from automated evidence alone; Plan 022
  requires the user-returned batched checkpoint.
- Preserve exact `System_Key` identity, captured-workbook binding, one modeless
  form/workbook per role, and headless Core/Domain authority.

## Evidence and traceability

- Reusable packaged Production plus clean restart: 2/2, with source contract
  6/6; see `tests/integration/plan022_slice4x_restart_results.md`.
- Dedicated NAS launcher readiness: initial 14/16 RED because Production reads
  saved the Designs workbook; final 16/16 GREEN after save-on-dirty-only and
  idempotent schema formatting; see
  `tests/integration/plan022_slice4x_nas_launcher_results.md`.
- Packaged XLAM/restart 74/74; packaged Ribbon/compile 142/142.
- Live role workflows 47/47; ordered Release 1 chain 30/30.
- Deterministic static baseline 19/19; reviewed cleanup 13/13; 152 components,
  4,987 procedures, 1,033 candidates, 8 literal `Application.Run` targets, 45
  unresolved dynamic calls, and 189 duplicate-body groups.
- The exact tested package set is committed in `deploy/current`.

## Do Not Repeat

- Do not count headless Inventory/Designs workbooks opened during startup as
  duplicate operator workbooks; count station-local role workbooks separately.
- Do not restore unconditional Designs workbook saves or repeated assignment
  of an already-correct text number format; both caused canonical hash changes
  from launcher-only reads.
- Do not repeat automated package/NAS suites unless binaries or relevant source
  change; all current gates are GREEN.

## Assumptions to Re-verify

- The dedicated NAS test warehouse and the existing human UAT credential remain
  available when the operator begins the checkpoint.
- Excel startup registration has not been changed since 2026-08-23.
- The visible checkpoint may expose layout or workflow defects that automation
  cannot judge; treat exact operator observations as new D13 RED if any fail.

## Open questions and blockers

- Blocked on the user performing and returning Plan 022 section 6 steps 1-17,
  including visible reusable Production workflow/layout evidence.
- No implementation, compile, regression, deployment, or NAS-readiness blocker
  remains.

## Immediate next action

The user runs Plan 022 section 6 steps 1-17 with the supplied validation CSV
and returns pass/fail, exact dialogs, workbook filenames, and redacted images.

## Critical references

- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, section 6
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, Production controls
- `tests/fixtures/plan022-demo-inventory-upload.csv`
- `tests/integration/plan022_slice4x_reusable_production_results.md`
- `tests/integration/plan022_slice4x_restart_results.md`
- `tests/integration/plan022_slice4x_nas_launcher_results.md`
- `src/Production/Forms/frmProduction.frm`
- `src/Production/Modules/modProductionReusableRun.bas`
- `src/DesignsDomain/Modules/modDesignsRuntime.bas`
- `src/DesignsDomain/Modules/modDesignsSchema.bas`
