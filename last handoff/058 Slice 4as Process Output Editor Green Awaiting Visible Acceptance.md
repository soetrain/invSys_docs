# Slice 4as Process Output Editor GREEN Awaiting Visible Acceptance

## Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11. Slice 4as removes
the Process Designer Output editor gap, places all visible Output fields on one
row, and replaces free-text Output UOM with the Settings catalog dropdown;
automated acceptance is GREEN and visible acceptance remains open.

## Current verified state

- Last verified: 2026-08-30.
- Code `main`: `fb30d54` (pushed).
- Documentation feature record `main`: `d2a618d` (pushed before this handoff).
- Active slice: Plan 022 Slice 4as -- Compact Process Output editor and catalog
  UOM.
- Final Core/Operations packages are rebuilt in `deploy/current`; Excel is
  closed.
- Preserved unrelated content: code `AGENTS.md`; unstaged NAS note in
  `invSys-Design-v4.11.md`; untracked `invSys dev notes.txt`.

## Decisions and constraints

- The correction is specifically Process Designer's Outputs editor.
- Visible order is ID / Output / Design / Ver / Output Qty / Yield % / Yield
  basis / UOM, all on one row.
- Hidden Output SKU state is retained in a nonvisible 1x1 control and reserves
  no visible gap.
- `cmbProcessOutputUom` is dropdown-list-only and reads the current warehouse
  Recipe UOM Catalog. Add/Update reject a missing/noncatalog selection.
- Selecting an existing Output restores its saved catalog UOM before Update.
- Designs Domain authority and persisted Output fields are unchanged.

## Evidence and traceability

- Focused Slice 4as: RED `1/6`, GREEN `6/6`.
- Prior Slice 4ar and Production layout: `8/8`, `8/8` GREEN.
- Packaged public Production actions/clean restart: `2/2` GREEN.
- Packaged XLAM `81/81`; Ribbon/VBA compile `142/142`; live roles `47/47`;
  ordered Release 1 `30/30`; dedicated NAS `16/16`; static `19/19`; reviewed
  growth `13/13`.
- Metrics: 154 components, 5,210 procedures, 1,048 candidates.
- Evidence: `tests/integration/plan022_slice4as_process_output_editor_green_results.md`.
- Packaged report: `reports/runtime/plan022-slice4as-final/production-reusable-production.md`.

## Do Not Repeat

- Do not restore a visible-width hidden Output SKU control.
- Do not use a free-text Output UOM or place UOM on the buttons row.
- An initial Release 1 attempt ended `28/29` only because of a temporary-path
  harness exception after all functional assertions passed. The clean final
  run is `30/30` and is the valid evidence.

## Assumptions to Re-verify

- The installed add-ins still resolve to the final `deploy/current` package
  hashes when the operator opens Excel.
- The saved Process remains available for selection and Update testing.

## Open questions and blockers

- Visible operator confirmation of the compact row and working UOM dropdown is
  unresolved.
- Broader Slice 4ar multi-Process Recipe/Run visible acceptance remains open.

## Immediate next action

Open packaged Production, select the saved Process in Process Designer, and
confirm the one-row Outputs editor and Recipe UOM Catalog dropdown visually.

## Critical references

- `src/Production/Forms/frmProduction.frm`.
- `tests/tooling/Test-Plan022Slice4asProcessOutputEditor.ps1`.
- `tools/validate_plan022_packaged_launchers.ps1`.
- `tests/integration/plan022_slice4as_process_output_editor_green_results.md`.
- `../invSys_docs/0 plan docs/xlam_invSys/invSys-Design-v4.11.md` D15.
- `../invSys_docs/expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4as.
- `../invSys_docs/0 plan docs/xlam_invSys/invSys-Controls-v1.md`, Slice 4as.
