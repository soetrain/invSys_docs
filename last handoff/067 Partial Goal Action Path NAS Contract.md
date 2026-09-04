# Partial Goal — Action Path NAS Contract

## Goal and release outcome

Continue Release 1 acceptance after accepted Production and D16 deployment by completing Slice 4bd physical Aggregator UAT and implementing the approved Slice 4be Action Path training library.

## Current verified state

- Code `main`: `75276d2`, pushed; docs `main` before this handoff: `29cc7f7`, pushed.
- NAS release `R1-20260904-75276d2` is applied to this station for next Excel startup.
- Slice 4bd source-set isolated D13 evidence is GREEN; physical two-computer/two-warehouse proof remains open.
- Production is user accepted; D16 NAS updater/rollback is physically accepted. GitHub is developer-only.
- Excel was closed at last deployment, but re-check before building/deploying.

## Decisions and constraints

- Aggregator uses current/remembered connected NAS roots automatically; Add Server is only for an additional NAS server. It never changes Send To.
- Action Paths are user-authored, versioned, non-executable training records under D18. They do not automate controls or rewrite event/inventory authority.
- User corrected a prior mistaken SharePoint assumption: the Action Path library is NAS-only at `<WarehouseRuntimeRoot>\Training\ActionPaths\<WarehouseId>`.
- `ACTION_PATH_MAINT` is required to save/export/import; signed-in Viewer users can search/read published paths.

## Evidence and traceability

- Slice 4bd tests: `Test-Slice4bdMultiServerSourceSet.ps1` (7/7 latest), `Test-Slice4bdAdminAggregatorCommand.ps1`, `Test-Slice4bdExplicitSourceAggregation.ps1`, `Test-Slice4bdAdminFormSmoke.ps1`.
- D18 storage contract: `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`, D18.

## Do Not Repeat

- Do not introduce a SharePoint dependency for Action Paths.
- Do not treat an Action Path as macro recording, control automation, or an audit assertion.
- Do not pass object types across XLAM boundaries.

## Open questions and blockers

- Slice 4be has no implementation yet. It needs D13 RED first for Viewer event selection/save/search/version warning and NAS library non-authority.
- The active product goal remains system-marked blocked; this handoff records its real partial completion because the goal API has no partial status.

## Immediate next action

Write and run the D13 RED test for public Viewer Action Path selection/save against the NAS training library, then implement the Viewer Action Path tab without changing current Events behavior.

## Critical references

- `src/Operations/Forms/frmInventoryViewer.frm`
- `src/Operations/Modules/modInventoryViewer.bas`
- `src/Core/Modules/modAuth.bas`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md` D18
