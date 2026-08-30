# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Slice 4ap now
makes the Recipe graph directional and name-visible while preserving exact
hidden graph identity and unconnected finished-output creation.

# Current verified state

- Code and docs are on `main`; Slice 4ap is pushed as code `e28b1f2` and docs
  `8e107a3`.
- Active slice: Plan 022 Slice 4ap, automated GREEN and awaiting visible UAT.
- Last verified 2026-08-30: Excel is closed; `deploy/current` contains the exact
  tested rebuilt five-package set.
- Preserve unrelated code `AGENTS.md`, the normative-spec NAS note, and
  untracked `invSys dev notes.txt` as user work.

# Decisions and constraints

- A connection reads **Upstream Process / Output / Downstream Process / Input
  Requirement / Qty / % / UOM**. Upstream and downstream nodes must differ;
  same-node/self connections are rejected.
- For the user's chain, the chai output from N3 connects to N4's chai input.
  N4 bottle and cap requirements resolve from assigned managed inventory.
- N4's bottled-product output remains unconnected; completion creates its
  unconsumed quantity as managed finished inventory with a new `System_Key`.
- Process, node, Output, and Requirement IDs remain hidden bound/persisted
  keys. Operator lists and selectors display business names.
- The raw seven-column `lstRecipeConnections` remains hidden staging authority;
  `lstRecipeConnectionDisplay` is the header-backed full-width name projection.

# Evidence and traceability

- Focused source: `2/9` RED, then `9/9` GREEN; Slice 4ao `7/7`, Slice 4an
  `7/7`, and Production layout `8/8` remain GREEN.
- Packaged Production action plus clean restart: `2/2`, including named
  selectors/projection, headers, full-width layout, visible selection,
  Connect/Update/Disconnect, same-node rejection, and final-output guidance.
- Packaged XLAM `81/81`; Ribbon/compile `142/142`; live roles `47/47`;
  ordered Release 1 `30/30`; dedicated NAS `16/16`; static `19/19`; growth
  `13/13`.
- Static metrics: 154 components, 5,186 procedures, 1,045 candidates; duplicate
  and dynamic-call ratchets did not grow.

# Do Not Repeat

- Do not connect N4's output back to N4 merely to represent the final product.
- Do not persist displayed names in place of node, Output, or Requirement IDs.
- Do not restore the side-by-side, headerless raw-ID Connections list.

# Assumptions to Re-verify

- The user's saved multi-Process Recipe draft remains available in the normal
  Production operator workbook.
- The user's N3 output and N4 requirement names still identify the intended
  chai-to-bottling edge after reopening.

# Open questions and blockers

- No automated blocker. Visible Recipe Designer connection, save, release, and
  final-output acceptance remain pending.

# Immediate next action

Open Recipe Designer, connect the named N3 chai output to N4's named chai input
requirement, leave N4's bottled-product output unconnected, then save/release
the Recipe and confirm the full-width named Connections row.

# Critical references

- `src/Production/Forms/frmProduction.frm`
- `tests/tooling/Test-Plan022Slice4apNamedRecipeGraph.ps1`
- `tools/validate_plan022_packaged_launchers.ps1:ProductionReusable`
- Architecture v4.11 D15 Recipe graph contract
- Plan 022 Slice 4ap
- `invSys-Controls-v1.md` Recipe Designer Connections row
