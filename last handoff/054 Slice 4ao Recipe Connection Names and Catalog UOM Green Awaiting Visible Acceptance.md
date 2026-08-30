# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Slice 4ao now
presents Recipe connection outputs by business name and restricts connection
UOM to the Settings catalog without changing persisted graph identity.

# Current verified state

- Code and docs are on `main`; Slice 4ao is pushed as code `662588d` and docs
  `3bd295e`.
- Active slice: Plan 022 Slice 4ao, automated GREEN and awaiting visible UAT.
- Last verified 2026-08-29: Excel is closed; `deploy/current` contains the exact
  tested rebuilt five-package set.
- Preserve unrelated code `AGENTS.md`, the normative-spec NAS note, and
  untracked `invSys dev notes.txt` as user work.

# Decisions and constraints

- `cmbConnectionOutput` displays `OutputName` but binds and persists the exact
  generated `OutputId`; names never replace graph identity.
- `cmbConnectionUom` is a non-free-text dropdown populated from the current
  Settings Recipe UOM Catalog.
- Connect and Update preserve the hidden output ID and selected catalog UOM.
- Designs Domain schemas, compatible-UOM validation, and run identity are
  unchanged.

# Evidence and traceability

- Focused source: `0/7` RED, then `7/7` GREEN; Slice 4an remains `7/7`.
- Packaged Production public action plus clean restart: `2/2`, including
  `RecipeOutputNameVisible=True`, `RecipeOutputIdPreserved=True`,
  `RecipeUomCatalog=True`, and `RecipeConnectionUpdated=True`.
- Packaged XLAM `81/81`; Ribbon/compile `142/142`; live roles `47/47`;
  ordered Release 1 `30/30`; NAS `16/16`; static `19/19`; growth `13/13`.
- Static metrics: 154 components, 5,180 procedures, 1,044 candidates; duplicate
  and dynamic-call ratchets did not grow.

# Do Not Repeat

- Do not persist the displayed output name in `FromOutputId`.
- Do not restore free-text Recipe connection UOM or create a second UOM catalog.
- Do not treat the generated output ID as the operator-facing label.

# Assumptions to Re-verify

- The user's saved multi-Process Recipe draft remains available in the normal
  Production operator workbook.

# Open questions and blockers

- No automated blocker. Visible Recipe Designer save/release acceptance remains
  pending.

# Immediate next action

Open Production Recipe Designer, select a source node, confirm Output shows its
name and UOM is a Settings-backed dropdown, then Connect and save/release the
multi-Process Recipe.

# Critical references

- `src/Production/Forms/frmProduction.frm`
- `tests/tooling/Test-Plan022Slice4aoRecipeConnectionEditor.ps1`
- `tools/validate_plan022_packaged_launchers.ps1:ProductionReusable`
- Architecture v4.11 D15 Recipe graph contract
- Plan 022 Slice 4ao
- `invSys-Controls-v1.md` Recipe Designer Connections row
