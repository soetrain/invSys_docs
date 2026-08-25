# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4ab fixes the visible Production item picker so current managed inventory can be
assigned to reusable Process requirements.

# Current verified state

- Last verified 2026-08-25: code `main` at `8e9e2ab`; docs `main` at `d502584`;
  both pushed to `origin/main`.
- Active slice: Plan 022 Slice 4ab, automated GREEN; the user's visible picker
  inventory/selection retest is pending.
- User visibly accepted Process-table creation, multiple tables, formula
  percentages, and automatic opening of Production Item Search.
- Rebuilt Core and Operations packages are in `deploy/current`; Excel is closed.
- User-owned uncommitted changes remain intentionally untouched: code
  `AGENTS.md` and the NAS/server note at the top of the normative specification.

# Decisions and constraints

- Process assignment search reads active exact Inventory Domain entities with
  nonblank `System_Key`, then deduplicates them to acceptable managed SKU/item
  alternatives.
- Selecting an alternative writes only item/SKU identity into the numbered
  Process pair. It does not allocate or persist the physical `System_Key`;
  Production Run - List retains exact-key allocation authority.
- The fix is Process-only in the Core-owned picker and does not restore or use
  prohibited legacy `ROW` identity.
- Captured-workbook binding, current-warehouse isolation, headless Domain
  authority, multi-table retrieval, and clean-restart behavior are preserved.

# Evidence and traceability

- Visible RED: picker opened but reported no managed inventory while Inventory
  Viewer showed seeded rows.
- Focused RED/GREEN:
  `tests/integration/plan022_slice4ab_process_picker_inventory_red_results.md`
  (`1/4`) and
  `tests/integration/plan022_slice4ab_process_picker_inventory_green_results.md`
  (`4/4`).
- Packaged public Production callback plus clean restart: `2/2`, recording
  `PickerOpened=True` and `PickerInventoryRows=True` through the actual worksheet
  selection event.
- Regressions: Slice 4aa `8/8`, Slice 4z `7/7`, packaged XLAM `74/74`,
  Ribbon/compile `142/142`, live role workflows `47/47`, Release 1 chain
  `30/30`, launcher contracts `24/24`, NAS `16/16`, static `19/19`, and reviewed
  cleanup `13/13`.

# Do Not Repeat

- Do not route the Process picker through `BuildInventoryPickerItemsFromTable`;
  that legacy loader still expects prohibited `ROW` and caused the empty list.
- Do not make opening the operator picker rebuild or mutate canonical inventory.
  The picker query is read-only; test projection preparation belongs in the
  isolated harness fixture.
- Do not persist the source `System_Key` into a Process acceptable-item pair;
  assignment is SKU-level and run allocation is exact-entity-level.
- If the full-chain runtime extractor reports only a temporary path after all
  business assertions pass, check for a harness-owned hidden Excel process and
  rerun from zero Excel processes; the clean-state rerun was `30/30`.

# Assumptions to Re-verify

- Confirm normal installed Excel loads the new Core/Operations package hashes,
  not cached pre-Slice 4ab add-ins.
- Confirm the user's seeded entities are active with positive available
  quantity; the picker intentionally lists currently available entities before
  SKU deduplication.

# Open questions and blockers

- No automated blocker remains. Visible confirmation that the user's seeded
  inventory appears and that choosing one fills the correct numbered
  acceptable item/SKU pair is unresolved.

# Immediate next action

Launch Production normally, enter **Acceptable Managed Item 1**, confirm seeded
inventory appears, choose one row, and verify the visible item plus matching
hidden **Accepted SKU 1** survive retrieval and Process reload.

# Critical references

- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md` Slice 4ab
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` Slice 4ab and section 8.2
- `src/Core/ClassModules/cDynItemSearch.cls`
- `src/Production/Forms/frmProduction.frm`
- `src/Production/Modules/mProduction.bas`
- `tests/tooling/Test-Plan022Slice4abProcessPickerInventory.ps1`
- `mProduction.BtnOpenProductionForm`
- `mProduction.RunProcessWorksheetBulkImportContractTest`
