# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Slice 4aq now
makes Recipe Designer route compatible outputs through visible parallel and
converging stages while preserving hidden exact graph identity.

# Current verified state

- Code and docs are on `main`; Slice 4aq is pushed as code `9db4703` and docs
  `5f5ae00`.
- Active slice: Plan 022 Slice 4aq, automated GREEN and awaiting visible UAT.
- Last verified 2026-08-30: Excel is closed; `deploy/current` contains the exact
  tested rebuilt five-package set.
- Preserve unrelated code `AGENTS.md`, the normative-spec NAS note, and
  untracked `invSys dev notes.txt` as user work.

# Decisions and constraints

- Recipe Designer is output-first: select **Produced by** and **Output**;
  **Feeds Process** lists only Recipe nodes with one unresolved input whose UOM
  and Ingredients Assignment item/SKU match that output.
- The matching `RequirementId` is internal identity, not an operator ingredient
  dropdown. Recipe Designer never defaults to the next Process in list order.
- **Output Flow** displays Stage / Produced by / Output / Feeds Process / Qty /
  % / UOM. Independent Processes share a Stage, outputs can converge, and an
  unconnected output displays **Finished inventory**.
- Direct graph edges express requirements: Brewed Black Slurry Tea and Chai
  Spice Mix feed Concentrate; Concentrate feeds Bottle; bottle/cap remain the
  Bottle Process's externally allocated inputs; bottled chai is terminal.
- Auto Order derives visible stages but persisted execution ordinals remain
  unique and deterministic. The seven-field Designs Domain edge schema is
  unchanged.

# Evidence and traceability

- Focused source: `2/8` initial presence and six behavioral REDs, then `8/8`
  GREEN; Slice 4ap `9/9`, 4ao `7/7`, 4an `7/7`, and Production layout `8/8`.
- Packaged Production action plus clean restart: `2/2`, including compatible
  targets, hidden requirement binding, no ingredient dropdown, fork/convergence
  Stages, terminal output, Connect/Update/Disconnect, Auto Order, lifecycle,
  Production Run, and restart persistence.
- Packaged XLAM `81/81`; Ribbon/compile `142/142`; live roles `47/47`;
  ordered Release 1 `30/30`; dedicated NAS `16/16`; static `19/19`; growth
  `13/13`.
- Static metrics: 154 components, 5,201 procedures, 1,048 candidates; duplicate
  and dynamic-call ratchets did not grow.
- One packaged attempt was interrupted by an AutoRecovered workbook/RPC failure
  before Recipe Designer; the clean rerun passed. One live-role attempt failed
  in unrelated Shipping hold setup with two Excel instances; its clean rerun
  passed `47/47`.

# Do Not Repeat

- Do not treat node list order as a connection or show a downstream Process's
  arbitrary ingredient list after selecting an output.
- Do not connect a Process to itself or connect the final bottled output back to
  its own Process.
- Do not persist displayed names in place of node, Output, or Requirement IDs.

# Assumptions to Re-verify

- The user's four released Process versions and their Ingredients Assignment
  alternatives remain available in the normal Production operator workbook.
- The concentrate Process requirements accept the brewed-tea and spice-mix
  output SKUs with matching UOMs, and the Bottle Process accepts the concentrate
  output SKU.

# Open questions and blockers

- No automated blocker. Visible construction, save, release, and Production Run
  acceptance of the user's four-Process sample remain pending.

# Immediate next action

Open Recipe Designer and confirm the named Output Flow is Brewed Tea + Chai
Spice Mix at Stage 1 -> Concentrate at Stage 2 -> Bottle at Stage 3 -> Finished
inventory, then save/release and load it in Production Run - List.

# Critical references

- `src/Production/Forms/frmProduction.frm`
- `tests/tooling/Test-Plan022Slice4aqOutputFlow.ps1`
- `tools/validate_plan022_packaged_launchers.ps1:ProductionReusable`
- Architecture v4.11 D15 Recipe output-flow contract
- Plan 022 Slice 4aq
- `invSys-Controls-v1.md` Recipe Designer Output routing / Output Flow rows
