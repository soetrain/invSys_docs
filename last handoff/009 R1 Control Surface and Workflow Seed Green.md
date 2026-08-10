# R1 Control Surface and Workflow Seed Green

## Goal and release outcome

Complete invSys Release 1 control-review answers 1-3 and 6-7 under Architecture
v4.11, leaving a tested surface and complete demo workflow kit for the user's
review of checklist items 4-5.

## Current verified state

- Last verified: 2026-08-09.
- Code: `main` at `852593a` (`Complete R1 control surface and demo workflow`),
  pushed to `origin/main`.
- Docs before this handoff: `main` at `2294e2f`; the current session's spec,
  Plan 022, controls catalog, and handoff changes are in the docs commit that
  contains this record.
- Active Plan 022 state: automated slices 4e (surface/Viewer), 4f (19-entity
  seed), and 4g (Shipping status anchor) are GREEN. Visible NAS checks and
  checklist items 4-5 remain open; slice 4d Shipping D14 `ROW` remains open.
- Excel is closed. The five packages in `deploy/current` were rebuilt and
  validated together.
- The user's pre-existing code `AGENTS.md` edit and normative-spec NAS pointer
  edit remain uncommitted and were not staged.

## Decisions and constraints

- The Release 1 form set contains 17 reviewed forms. Eighteen unreachable or
  empty form shells and unconstructed controls were removed with protecting
  regression evidence.
- Item search uses the single Core `frmItemSearch` runtime canvas with role
  profiles. Do not restore role-named search shells or dynamic template copies.
- Operations exposes a signed-in, capability-neutral Inventory Viewer. It reads
  only the current published snapshot on explicit refresh and never mutates an
  authority workbook.
- Purchasing remains an intentional visible stub inside Receiving; no separate
  launcher exists.
- Admin Seed Demo Inventory creates 19 new entities per action. Repeated seed is
  intentional new-entity creation, not SKU upsert or legacy migration.
- Production, Viewer, and Shipping must preserve readable anchored proportions.
  Shipping status stays Left + Top + Right with fixed height.
- Do not correct Shipping `ROW` silently. Checklist item 5 and Plan 022 slice 4d
  require the focused public form-action RED and a `System_Key` correction.

## Evidence and traceability

- Focused GREEN: R1 surface 6/6; demo seed 4/4; Shipping anchor 4/4.
- Packaged Viewer: three levels loaded, repeated launch reused one form,
  filtering returned one row, and snapshot hash remained unchanged.
- Packaged Admin seed: 19 unique `System_Key` entities, all `GOOD`, with raw,
  WIP, shippable, and packaging coverage through canonical inventory, catalog,
  snapshot, saved Receiving projection, and actual Receiving Refresh.
- Create Warehouse integration: 15/15; repeated seed added 38 unique entities.
- Ordered packaged Release 1 chain: 30/30 through Receiving, two Production
  batches, Boxing, Shipping, restart, and reconciliation.
- Repeated packaged launchers: Receiving, Production, and Shipping 3/3.
- Packaged XLAM 74/74; RibbonX 136/136; ribbon generation 46/46; Operations
  shadow 13/13 with zero unresolved collisions.
- Static maintenance baseline 19/19: 149 components, 4,566 procedures, 955
  scanner candidates, 9 literal `Application.Run` targets, and 48 unresolved
  dynamic calls.

## Do Not Repeat

- Do not restore removed shells merely because scanner output names old source;
  package, source, collision, and full-chain regressions protect the reviewed set.
- Do not use Production DPI-derived form zoom; packaged geometry is GREEN at
  `Zoom=100`.
- Do not bottom-anchor Shipping `txtStatus`.
- Do not seed only three transport-proof rows or reuse the sanitized catalog's
  legacy mutation/`ROW` identity path.
- Do not commit or overwrite the user's existing `AGENTS.md` or normative-spec
  NAS pointer edits.

## Assumptions to Re-verify

- The dedicated NAS warehouse and sign-in remain available for visible UAT.
- The next Excel session loads the five package hashes from code commit
  `852593a`.
- Native maximize/restore behavior on the user's display matches packaged
  geometry tests.

## Open questions and blockers

- Unresolved visible acceptance: all 19 seeded entities after Admin Seed and
  Receiving Refresh on the dedicated NAS warehouse.
- Unresolved visible acceptance: Inventory Viewer launch/filter/resize,
  Production maximize/restore, and Shipping grow/shrink/maximize/restore.
- Unresolved review: checklist items 4 (action effect clarity) and 5 (durable
  identity); Shipping's visible/backing `ROW` path is the known D14 conflict.

## Immediate next action

Read `invSys-Controls-v1.md` and answer checklist items 4-5, then run the single
batched visible acceptance checkpoint against the rebuilt five-package set.

## Critical references

- `../0 plan docs/xlam_invSys/invSys-Design-v4.11.md`
- `../0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `../expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
- `src/Core/Modules/modInventoryViewerData.bas`
- `src/Operations/Forms/frmInventoryViewer.frm`
- `src/Admin/Modules/modAdminInventorySeed.bas`
- `src/Shipping/Forms/frmShipmentsTally.frm`
- `tools/validate_inventory_viewer.ps1`
- `tools/validate_admin_seed_inventory_callback.ps1`
- `tools/validate_release1_full_chain.ps1`
- `tests/integration/inventory_viewer_results.md`
- `tests/integration/admin_seed_callback_green_results.md`
- `tests/integration/slice14_results.md`
- `deploy/current/addins-manifest.json`
