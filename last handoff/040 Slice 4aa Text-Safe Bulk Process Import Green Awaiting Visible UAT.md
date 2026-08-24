# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4aa now makes the Process worksheet text-safe, catalog-backed, searchable, and
capable of Ctrl+click multi-table DRAFT import, advancing D15 visible Production
acceptance.

# Current verified state

- Last verified 2026-08-24: code `main` at `90e075c`; docs `main` at `94b085c`;
  both pushed to `origin/main`.
- Active slice: Plan 022 Slice 4aa, automated GREEN; visible operator UAT is
  pending. Production Run - Tree remains experimental/out of scope.
- Deployed Core, Designs Domain, and Operations packages were rebuilt in
  `deploy/current`; Excel is closed.
- User-owned uncommitted changes remain intentionally untouched: code
  `AGENTS.md` and the NAS/server note at the top of the normative specification.

# Decisions and constraints

- Process, Requirement, Output, and generated Design IDs are automatic locked
  three-character Base-36 text values; numeric-only IDs are canonicalized at
  processor ingress so Excel cannot turn `001` into `1`.
- INPUT rows own automatic Requirement IDs and four initial numbered acceptable
  managed-item/hidden-SKU pairs. **Add Acceptable Item** appends more pairs.
- Entering a numbered acceptable-item cell by mouse, Tab, or Enter opens the
  existing Core item picker and writes to that numbered pair.
- UOM validation comes from Settings' Recipe UOM Catalog.
- **Retrieve Selected Process** resolves every selected Process table across
  Ctrl+click areas, validates all before writes, saves through the public
  Process DRAFT action, deletes successful selected tables only, and preserves
  failed/unselected tables. Release/obsolete remain explicit actions.
- Historical vertical `ALTERNATIVE` import, captured-workbook binding, headless
  Core/Domain authority, and packaged launcher reuse remain preserved.

# Evidence and traceability

- Focused RED: `tests/integration/plan022_slice4aa_process_bulk_import_red_results.md`
  (`1/8`, seven behavioral REDs).
- Focused GREEN: `tests/integration/plan022_slice4aa_process_bulk_import_green_results.md`
  (`8/8`); historical Slice 4z source remains `7/7`.
- Packaged public Production callback plus clean restart: `2/2`; actual picker
  opened and all eight Slice 4aa evidence fields are true.
- Regressions: packaged XLAM `74/74`, Ribbon/compile `142/142`, role workflows
  `47/47`, Release 1 chain `30/30`, launcher contracts `24/24`, NAS `16/16`,
  static `19/19`, reviewed cleanup `13/13`, and layout GREEN for 3 sizes x 5
  pages plus native window transitions.
- Accepted formula behavior remains 611.2 lb and 16.4/32.7/1.8/49.1%, totaling
  100.0%.

# Do Not Repeat

- Do not format Requirement/Design formula columns as text; Excel then stores
  formulas literally. Identity value columns/metadata stay text, while formula
  columns remain General.
- Do not treat automatic formula/ID cells as business data when deciding
  whether an unused template row is populated.
- Formatting only the event writer or Designs table is insufficient: local
  staging/inbox reads can coerce numeric-only Base-36 IDs, so keep the processor
  ingress canonicalization.
- Do not require all worksheet tables to disappear after bulk retrieval;
  selected successful tables disappear and unselected tables remain.

# Assumptions to Re-verify

- Confirm normal installed Excel loads the newly deployed package hashes rather
  than an older cached add-in set.
- Confirm the user's current Recipe UOM Catalog contains every UOM used in the
  visible Process tables.

# Open questions and blockers

- No automated blocker remains. Operator-visible acceptance of generated IDs,
  picker behavior, added alternatives, Ctrl+click import, and subsequent Recipe
  Designer ordering is still unresolved.

# Immediate next action

Open the packaged Production launcher and perform Plan 022 UAT step 15 against
two Ctrl+click-selected Process tables, beginning by confirming `001`-style IDs,
the UOM dropdown, actual item-search entry, and selected-only retrieval.

# Critical references

- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md` D15
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md` Slice 4aa and UAT step 15
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` Slice 4aa and section 8.2
- `src/Production/Modules/modProductionProcessWorksheet.bas`
- `src/Production/Forms/frmProduction.frm`
- `src/Core/ClassModules/cDynItemSearch.cls`
- `src/Core/Modules/modProcessor.bas`
- `mProduction.BtnOpenProductionForm`
- `mProduction.RunProcessWorksheetBulkImportContractTest`
