# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Slice 4am now
allows active zero-quantity managed-item creation without treating zero stock
as allocatable Production inventory.

# Current verified state

- Code: `main` at `878ef3b` before the Slice 4am commit; docs: `main` at
  `186a06b` before the Slice 4am commit. Both will be committed and pushed as
  the final action of this session.
- Active slice: Plan 022 Slice 4am, automated GREEN and awaiting visible UAT.
- Excel may be closed for rebuild/deployment under standing user permission.
- Preserve the user's unrelated code `AGENTS.md` change and normative-spec NAS
  note as uncommitted work.

# Decisions and constraints

- Counted Add Item and worksheet `ADD` accept explicit numeric quantity `0` and
  reject blank, nonnumeric, and negative quantities.
- Zero creation owns one immutable active `System_Key`; it is not retirement.
- Active zero definitions remain visible in managed inventory, Viewer,
  Receiving choices, and Process Designer's catalog picker.
- Production Run exact available allocation remains positive-only. Later
  receipt or Production output creates a new physical exact key.

# Evidence and traceability

- Focused D13: initial `1/8` RED; strengthened rebuild/picker checks exposed
  intermediate RED; final `8/8` GREEN.
- Packaged action `81/81`; Ribbon/compile `142/142`; live roles `47/47`;
  ordered Release 1 `30/30`; dedicated NAS runtime `16/16`.
- Deterministic static `19/19`; reviewed growth/cleanup `13/13`.
- Governing records: Architecture v4.11 D14, Plan 022 Slice 4am, controls v1.

# Do Not Repeat

- Do not expose zero counted entities through `ListAvailableInventoryEntities`:
  that lets Production Run allocate stock that does not exist. Process Designer
  must use the managed catalog picker projection instead.
- Do not represent zero quantity as `DEPLETED` or catalog-only state; both hide
  a valid managed definition from required operator surfaces.

# Assumptions to Re-verify

- Visible packaged workbook behavior is not yet operator-confirmed after this
  deployment.

# Open questions and blockers

- No automated blocker. Visible UAT remains pending.

# Immediate next action

Create a disposable counted item with Starting Qty `0`, Refresh, and confirm it
appears in Add/Edit managed inventory, Viewer, Receiving, and Process Designer
while Production Run cannot allocate it.

# Critical references

- `src/Admin/Forms/frmAddInventoryItem.frm`
- `src/Admin/Modules/modAdminInventoryWorksheet.bas`
- `src/InventoryDomain/Modules/modInventoryApply.bas`
- `src/InventoryDomain/Modules/modInventoryQueries.bas`
- `src/Core/ClassModules/cDynItemSearch.cls`
- `src/Core/Modules/modInventoryViewerData.bas`
- `tests/tooling/Test-Plan022Slice4amZeroStartingQuantity.ps1`
