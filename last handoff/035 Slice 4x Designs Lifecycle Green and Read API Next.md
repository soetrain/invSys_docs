# Slice 4x Designs Lifecycle Green and Read API Next

## Goal and release outcome

Complete Plan 022 Slice 4x reusable Production Processes/Recipe graphs under
Architecture v4.11 D15, then resume saved-workbook/NAS Production acceptance
without regressing the completed Operations workflows.

## Current verified state

- Last verified: 2026-08-23.
- Code: `main` at `99a2bf8` (pushed to `origin/main`).
- Docs: `main` at `e9675f6` (pushed to `origin/main`).
- Active slice: Plan 022 Slice 4x, reusable Production Processes and Recipe
  graphs.
- Implemented first Designs Domain boundary: 13-table schema; generic
  definition event identity; Process/Recipe save, release, obsolete, replay,
  immutable versions, multi-output projection, ingredient alternatives,
  Recipe graph validation, and released-Recipe dependency protection.
- Published Core, Designs Domain, and Operations XLAMs match the committed
  source/manifest. Excel is closed.
- Preserved uncommitted user changes: code `AGENTS.md`; docs
  `0 plan docs/xlam_invSys/invSys-Design-v4.11.md` NAS/server note.

## Decisions and constraints

- D15 and Plan 022 Slice 4x remain authoritative. The current four-page
  Production form is intentional RED until distinct Process Designer and
  Recipe Designer operator handlers exist.
- `mProduction.BtnOpenProductionForm` remains the public packaged entry and
  captured station-local workbook binding remains mandatory.
- Designs Domain remains headless authority. Operations edits/queues events;
  it must not become canonical definition authority.
- The current Domain implementation is an incremental GREEN boundary, not
  Slice 4x completion. Read APIs, packaged form actions, run planning,
  multi-output inventory events, static ratchets, and visible UAT remain open.

## Evidence and traceability

- Packaged launcher/live-form RED: 0/1. Launcher, saved workbook reuse, and
  `0.001%`/`100%`/`1000%` scaling stayed GREEN; live form reported four pages,
  missing Process/Recipe Designer, and legacy Recipe Builder present.
- Initial Domain RED -> GREEN: 0/4 -> 4/4, range 264-267.
- Lifecycle/graph RED -> GREEN: 2/8 -> 8/8, range 268-275.
- Combined reusable range: 12/12 GREEN, range 264-275 of 309.
- Adjacent Designs/Core/Inventory range: 33/34. The sole failure,
  `TestInventoryQueries_PickerPublishesEverySkuLocation`, was already the sole
  failure in the 2026-08-17 baseline and neither its test nor query changed.
- Published five-XLAM compile/load/Ribbon validation: 74/74 GREEN.
- Static scan: candidate count 968 versus baseline 965; literal
  `Application.Run` 8/8, unresolved dynamic calls 47/47, and duplicate bodies
  185/185. New open findings are the `modDesignsApply` module/procedure size
  ratchets; no exception has been accepted.

## Do Not Repeat

- A raw line-range extraction intended to split `modDesignsApply` crossed VBA
  procedure boundaries and caused Process replay to stall. The exact working
  module was restored from the already-built 8/8-GREEN
  `invSys.Designs.Domain.xlam`, then the combined range returned to 12/12.
  Split by whole named procedures with a focused replay test after each move.
- Do not treat the pre-existing picker failure as a Slice 4x regression or
  rewrite its prohibited legacy `ROW` assertion inside this Production slice.

## Assumptions to Re-verify

- Re-run the static scan after the safe module split; the recorded 968/965
  comparison predates any future refactor.
- Rebuild the Designs Domain after subsequent source changes only while Excel
  is closed.
- Exact runtime control IDs for Process Designer and Recipe Designer are still
  intentionally unsettled until their form-action RED is added.

## Open questions and blockers

- Open gate: split reusable validation/replay from `modDesignsApply` without
  increasing bloat or dynamic-call metrics.
- Open implementation: released Process/Recipe read APIs and primitive bridge
  envelopes.
- Open D13 ranges: packaged Process/Recipe form handlers, two-batch List run,
  exact-key multi-output processor/application, snapshot, and Viewer Events.
- Open visible acceptance: dedicated NAS Production workflow and five-page
  maximize/restore polish.

## Immediate next action

Add and observe a focused RED for released Process/Recipe list/detail/graph
read APIs through the Designs Domain bridge before implementing those APIs.

## Critical references

- `src/DesignsDomain/Modules/modDesignsSchema.bas`
- `src/DesignsDomain/Modules/modDesignsApply.bas`
- `tests/unit/TestDesignsDomain.bas`
- `tools/run_phase6_excel_validation.ps1`
- `src/Production/Modules/mProduction.bas`
- `src/Production/Forms/frmProduction.frm`
- `tools/validate_plan022_packaged_launchers.ps1`
- `tests/integration/plan022_slice4x_*_results.md`
- `reports/runtime/plan022-slice4x-static/`
- Architecture v4.11 D15; Plan 022 Slice 4x; Controls catalog sections 8.1-8.5
