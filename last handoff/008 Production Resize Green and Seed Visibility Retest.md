# Production Resize Green and Seed Visibility Retest

## Goal and release outcome

Complete invSys Release 1 acceptance under Architecture v4.11 by closing the
remaining Plan 022 visible Seed/Receiving and Production checks, then correcting
Shipping's D14 `ROW` identity conflict test-first.

## Current verified state

- Last verified: 2026-08-04.
- Code: `main` at `c25e8a1` (`Fix Production scaling and strengthen seed
  validation`).
- Docs before this handoff: `main` at `3b6f4b6` (`Record Plan 022 operator UAT
  blocker`).
- Active Plan 022 slices: 4b Seed round trip is automated GREEN with visible
  Receiving retest pending; 4c Production resize is corrected and automated
  GREEN with visible retest pending; 4d Shipping D14 `ROW` remains open.
- The user's pre-existing code-repository `AGENTS.md` edit remains uncommitted
  and was not staged.
- Excel is closed. The five packages in `deploy/current` were rebuilt,
  registered, and validated together.

## Decisions and constraints

- The original operator Seed checkpoint remains RED even though later evidence
  proves the demo rows exist. Do not rewrite human acceptance from automation.
- Raw demo inventory visibility is checked in Receiving's top Inventory list
  after Refresh with filter `DEMO-`; Production Recipe Builder and Shipping's
  shippable list are not raw inventory views.
- Repeated seed actions intentionally create new immutable `System_Key` values;
  do not seed again merely to locate existing demo rows.
- Production must remain at `Zoom=100`; DPI-derived UserForm zoom reproduced the
  shrunken-controls defect.
- Do not silently fix Shipping `ROW` as part of unrelated acceptance work. D14
  requires a focused form-action RED and exact `System_Key` correction.

## Evidence and traceability

- Production focused RED: packaged maximize geometry reported `Zoom=60` and
  reproduced the small upper-left control surface.
- Production GREEN: 3 sizes x 4 pages, zero out-of-bounds controls, zero
  interactive overlaps, native minimize/restore/maximize/restore, and maximized
  client fill PASS; see `tests/integration/slice9_layout_results.md`.
- Seed/Receiving GREEN: 3 unique keys with `Condition=GOOD` match canonical
  inventory, published snapshot, saved Receiving workbook, and 3 filtered rows
  through the actual Refresh handler; see
  `tests/integration/admin_seed_callback_green_results.md`.
- Packaged XLAM: 54/54; RibbonX: 136/136; Release 1 full chain: 30/30.
- Static evidence: 164 components, 4,562 procedures, 964 scanner candidates,
  9 literal `Application.Run` targets, and 48 unresolved dynamic calls.
- Full control catalog:
  `../0 plan docs/xlam_invSys/invSys-Controls-v1.md`.

## Do Not Repeat

- Do not diagnose the Seed action as missing canonical or operator data without
  first checking the Receiving inventory control; saved projections already
  contained the demo entities.
- Do not restore `GetDpiForWindow`-derived form zoom in Production.
- Do not commit or overwrite the user's existing `AGENTS.md` modification.
- Do not save screenshots containing runtime roots, usernames, or RunIds.

## Assumptions to Re-verify

- The account startup registry still contains only the consolidated Operations
  and Admin leaf add-ins from `deploy/current`.
- The dedicated test warehouse and prepared sign-in remain available.
- The operator's next Excel session loads manifest hashes from code commit
  `c25e8a1`.

## Open questions and blockers

- Unresolved: whether existing demo rows are visible through the specified
  Receiving Refresh/filter controls in the user's normal Excel session.
- Unresolved: whether the corrected Production controls visibly fill and
  restore with the native window on the user's display.
- Blocker: Shipping still exposes/depends on prohibited `ROW` controls and has
  not received its required D13 focused correction.

## Immediate next action

Restart Excel, verify existing `DEMO-` rows in Receiving's top Inventory list,
then maximize and restore Production and record the visible results without
running Seed again.

## Critical references

- `../0 plan docs/xlam_invSys/invSys-Design-v4.11.md`
- `../0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `../expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
- `src/Production/Modules/modProductionFormWindow.bas`
- `src/Receiving/Forms/frmReceiving.frm`
- `src/Receiving/Modules/modTS_Received.bas`
- `tools/validate_slice9_production_layout.ps1`
- `tools/validate_admin_seed_inventory_callback.ps1`
- `tests/integration/plan022_results.md`
- `tests/integration/slice9_layout_results.md`
- `tests/integration/admin_seed_callback_green_results.md`
- `deploy/current/addins-manifest.json`
