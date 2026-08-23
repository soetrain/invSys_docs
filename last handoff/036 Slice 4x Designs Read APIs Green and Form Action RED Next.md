# Slice 4x Designs Read APIs Green and Form Action RED Next

## Goal and release outcome

Complete Plan 022 Slice 4x reusable Production Processes/Recipe graphs under
Architecture v4.11 D15, then resume saved-workbook/NAS Production acceptance
without regressing the completed Operations workflows.

## Current verified state

- Last verified: 2026-08-23.
- Code: `main` at `ae19cec` (pushed to `origin/main`).
- Docs: `main` at `369fec8` (pushed to `origin/main`).
- Active slice: Plan 022 Slice 4x, reusable Production Processes and Recipe
  graphs.
- Designs Domain Process/Recipe schema, lifecycle, replay, validation, and read
  APIs are implemented. Core and Operations expose the read APIs through one
  declared cross-XLAM query dispatcher and direct typed role wrappers.
- Published normative five-XLAM package set matches the committed source and
  manifest. Excel is closed.
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
- Exact-version Process/Recipe details cross the XLAM boundary as JSON strings;
  list results contain primitive arrays; validation uses a tab-delimited
  primitive status envelope.

## Evidence and traceability

- Designs Domain read API RED -> GREEN: 0/3 -> 3/3, range 276-278.
- Core cross-XLAM bridge RED -> GREEN: 0/1 -> 1/1, test 279.
- Operations primitive wrapper RED -> GREEN: 0/1 -> 1/1, test 280.
- Reusable Production range: 17/17 GREEN, range 264-280 of 314.
- Adjacent Designs/Core/Inventory range: 38/39, range 262-300. The sole
  failure is the pre-existing prohibited-`ROW` picker assertion recorded in the
  2026-08-17 baseline.
- Published five-XLAM compile/load/Ribbon validation: 74/74 GREEN.
- Clean-state packaged launcher check preserved second-launch reuse and
  `.001%`/`100%`/`1000%` scaling; expected surface result remains 0/1 because
  the form has four pages and the legacy Recipe Builder.
- Static scan: 967 candidates versus baseline 968; literal `Application.Run`
  8/8; duplicate-body groups 185/185; unresolved dynamic calls improved
  47 -> 45. Evidence is
  `tests/integration/plan022_slice4x_designs_read_api_results.md`.

## Do Not Repeat

- Do not split `modDesignsApply` by raw line ranges. A prior extraction crossed
  procedure boundaries and caused replay to stall. Move only whole named
  procedures with focused replay tests after each move.
- A packaged launcher retry hung when an Excel process from the preceding
  package validator remained alive. Verify zero Excel processes before launcher
  validation; the clean retry completed with the expected surface RED.
- Do not treat the pre-existing picker failure as a Slice 4x regression or
  restore its prohibited legacy `ROW` assertion path.

## Assumptions to Re-verify

- Re-run static maintenance after any form/controller addition; the current
  967/185/8/45 counts are the latest accepted comparison.
- Rebuild only while Excel is closed, and rebuild Core plus all reference
  consumers before packaged validation.
- Exact runtime control IDs for Process Designer and Recipe Designer remain
  intentionally unsettled until the focused handler/layout RED is recorded.

## Open questions and blockers

- Open gate: split reusable validation/replay from `modDesignsApply` by whole
  procedures without increasing bloat or dynamic-call metrics.
- Open D13 ranges: packaged Process/Recipe form handlers, ingredient assignment,
  two-batch List run, exact-key multi-output processor/application, snapshot,
  and Viewer Events.
- Open visible acceptance: dedicated NAS Production workflow and five-page
  maximize/restore polish.

## Immediate next action

Add and observe a focused packaged RED through `mProduction.BtnOpenProductionForm`
and the actual Process/Recipe form-action handlers before replacing the Recipe
Builder page.

## Critical references

- `src/DesignsDomain/Modules/modDesignsQueries.bas`
- `src/DesignsDomain/Modules/modDesignsBridgeApi.bas`
- `src/Core/Modules/modDesignsDomainBridge.bas`
- `src/Core/Modules/modOperationsPrimitiveBridge.bas`
- `tests/unit/TestDesignsDomain.bas`
- `tests/integration/plan022_slice4x_designs_read_api_results.md`
- `src/Production/Modules/mProduction.bas`
- `src/Production/Forms/frmProduction.frm`
- `tools/validate_plan022_packaged_launchers.ps1`
- Architecture v4.11 D15; Plan 022 Slice 4x; Controls catalog sections 8.1-8.5
