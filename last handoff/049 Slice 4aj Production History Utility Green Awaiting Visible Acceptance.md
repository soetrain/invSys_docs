# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4aj advances D15 by retaining completed Production batch rows, adding Used
Goods and Process Total, and projecting catalog Utility mode visibly without
changing exact-key inventory authority.

# Current verified state

- Last verified: 2026-08-27. Code `main` is at `3b92aa1`; docs `main` is at
  `e761218` before this handoff/documentation commit.
- Slice 4aj automated gates are complete; visible operator acceptance is open.
- The five rebuilt add-ins and current manifest are committed and pushed.
- Excel is closed before handoff.
- Intentionally uncommitted user changes remain in code `AGENTS.md` and the
  NAS/server note at the top of docs `invSys-Design-v4.11.md`.

# Decisions and constraints

- Production Output retains every completed Process-output row for the loaded
  run; Next Batch adds separate active rows instead of overwriting history.
- The nine visible columns are Process, Output, UOM, Last Actual, Batch, Used
  Goods, Process Total, Recall, and `System_Key`.
- Used Goods is the scaled compatible input quantity consumed by that Process
  for the batch. Process Total sums actual output only for the same
  Process/output/UOM.
- Historical rows are display-only. Actual Output selection/change handling
  maps only active-batch rows to output definitions.
- Inventory Domain exact-entity projections append catalog `TRACK_QTY`,
  `ITEM_KIND`, and Category while preserving the original ten envelope fields.
  Utility items display **Utility** in Run palette Inv and Inventory Check;
  required/used quantities remain measured and the existing non-counted apply
  rule remains authoritative.

# Evidence and traceability

- Focused RED: `1/7`; focused GREEN: `7/7`.
- Packaged Production plus clean restart: `2/2`, including
  `BatchHistoryRows=True`, `ProcessTotal=True`, and `UtilityDisplay=True` through
  the real public handlers.
- Prior focused regressions: Slice 4ag `6/6`, 4ah `5/5`, 4ai `8/8`; launcher
  source contracts `24/24`.
- Packaged XLAM `76/76`; Ribbon `142/142`; live roles `47/47`; ordered Release
  1 chain `30/30`; dedicated NAS runtime `16/16`; deterministic static `19/19`;
  reviewed growth `13/13`.
- Static inventory: 154 components, 5,151 procedures, 1,040 candidates. Full
  evidence is in
  `tests/integration/plan022_slice4aj_production_history_utility_green_results.md`.

# Do Not Repeat

- The reusable packaged workflow normally takes about four minutes because it
  creates/releases designs, completes two persisted batches, and clean-restarts
  Excel. High Excel CPU during that bounded run was not an infinite loop.
- One clean packaged attempt failed with Excel RPC `0x800706BE`; a fresh Excel
  retry was `2/2` GREEN.
- The first NAS attempt attached read-only extraction to leftover isolated-test
  Excel instances and failed package-hash observation. Closing every Excel
  process before retry produced `16/16` GREEN; do not treat the contaminated
  report as current evidence.
- Static baseline generation must use timestamp
  `2026-08-16T20:00:00Z` to satisfy the deterministic Slice 3 comparison.

# Assumptions to Re-verify

- The saved Production operator workbook reopens against the rebuilt package
  set after Excel restart.
- Filtered Water still has catalog Qty mode Utility and remains an acceptable
  input assignment for the tea Process.

# Open questions and blockers

- Visible Slice 4aj acceptance is unresolved: confirm multiple completed batch
  rows, the cumulative Process Total, and **Utility** instead of `19400`.
- Visible Slice 4ai Admin inventory worksheet acceptance remains unresolved.
- Production Run - Tree remains experimental and outside current acceptance.

# Immediate next action

Restart Excel, reopen Production Run - List, load the released tea Recipe,
confirm Filtered Water Inv shows **Utility**, complete or advance another batch,
and verify that prior batch rows remain with the expected Process Total.

# Critical references

- `src/Production/Forms/frmProduction.frm`
- `src/Production/Modules/modProductionReusableRun.bas`
- `src/InventoryDomain/Modules/modInventoryQueries.bas`
- `tests/tooling/Test-Plan022Slice4ajProductionHistoryUtility.ps1`
- `tools/validate_plan022_packaged_launchers.ps1` state `ProductionReusable`
- Architecture v4.11 D15 reusable Production Run contract
- Plan 022 **Slice 4aj -- Production batch history and Utility projection**
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` version 1.28
