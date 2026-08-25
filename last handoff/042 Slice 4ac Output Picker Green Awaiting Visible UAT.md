# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4ac makes every worksheet-defined Process OUTPUT select and retain a managed
item/SKU before visible Production acceptance continues.

# Current verified state

- Last verified 2026-08-25: code `main` at `e11bb6c`; normative/plan/control
  docs commit `a16ec75`; both are pushed to `origin/main` and this handoff is
  the current docs pointer.
- Active slice: Plan 022 Slice 4ac, automated GREEN; visible OUTPUT picker and
  retrieval UAT is pending.
- User visibly accepted that the Process picker now contains current managed
  inventory. INPUT pair selection remains part of final workflow acceptance.
- Rebuilt Core and Operations packages are in `deploy/current`; Excel is closed.
- Preserved uncommitted user work: code `AGENTS.md` and the NAS/server note at
  the top of Architecture v4.11.

# Decisions and constraints

- Entering an OUTPUT **Name** cell by mouse, Tab, or Enter opens the same
  Core-owned managed-item picker used by INPUT acceptable-item cells.
- Picker commit writes the visible output name and hidden system-managed
  **Output SKU**. It writes no source `System_Key`; completion still creates a
  new exact key for each output entity.
- Output ID and Design ID/version remain generated. INPUT numbered alternatives
  and Ingredients Assignment remain unchanged.
- Captured-workbook binding, current-warehouse isolation, headless Domain
  authority, packaged launcher reuse, and multi-table retrieval are preserved.

# Evidence and traceability

- Focused RED: `1/6`; OUTPUT Name was not a picker target, Core commit only
  understood INPUT pairs, and no OUTPUT SKU packaged proof existed.
- Focused GREEN: Slice 4ac `6/6`; prior 4ab `4/4`, 4aa `8/8`, 4z `7/7`.
- Packaged public Production callback plus clean restart: `2/2`, recording
  `OutputPickerOpened=True`, `OutputPickerCommitted=True`,
  `OutputSkuHidden=True`, `OutputSkuRoundTrip=True`, and `NoPhysicalKey=True`.
- Regressions: packaged XLAM `74/74`, Ribbon/compile `142/142`, live role
  workflows `47/47`, clean ordered Release 1 chain `30/30`, launcher contracts
  `24/24`, dedicated NAS `16/16`, static `19/19`, reviewed growth `13/13`.

# Do Not Repeat

- Do not store the selected inventory entity's `System_Key` in a Process
  definition; OUTPUT picker selection is managed SKU identity only.
- Do not put OUTPUT identity into a numbered INPUT alternative pair. OUTPUT
  selection belongs to Name plus hidden Output SKU.
- The first full-chain run passed all business assertions but hit the known
  temporary-path cleanup exception. After closing only the verified test-owned
  `/automation -Embedding` Excel process, the clean rerun passed `30/30`.

# Assumptions to Re-verify

- Confirm installed Excel loads the new Core/Operations hashes rather than a
  cached pre-Slice 4ac package.
- Confirm the chosen output item/SKU is the one the operator intends to produce;
  picker ordering is a projection and carries no physical-entity allocation.

# Open questions and blockers

- No automated blocker remains. Visible OUTPUT search, selection, and retrieval
  in the user's saved Production workbook is unresolved.

# Immediate next action

In an OUTPUT row, enter the **Name** cell, choose a managed item, retrieve that
Process table, and confirm the Process remains available with the selected
output identity.

# Critical references

- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md` Slice 4ac
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md` D15 Process contract
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` Slice 4ac and section 8.2
- `src/Production/Modules/modProductionProcessWorksheet.bas`
- `src/Core/ClassModules/cDynItemSearch.cls`
- `src/Production/Forms/frmProduction.frm`
- `tests/tooling/Test-Plan022Slice4acProcessOutputPicker.ps1`
- `mProduction.RunProcessWorksheetOutputPickerContractTest`
