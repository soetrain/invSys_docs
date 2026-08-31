# Goal and release outcome

Complete invSys Release 1 user acceptance under Architecture v4.11, beginning
with one uninterrupted multi-Process Production run and correct operator
visibility of every routed intermediate consumed by the final Process.

# Current verified state

- Code `main`: `fbdc2ee` (`Enable released Process editing and sheet export`).
- Docs `main`: `12495d0` (`Record Slice 4aw Process edit export acceptance`)
  before this handoff/plan commit.
- Active slice: Plan 022 Slice 4ax, newly recorded; no implementation has begun.
- Slice 4aw automated gates are GREEN; visible editing/releasing of the corrected
  bottling Process and full Recipe completion remain open.
- Last verified 2026-08-30: Excel is open and must be closed before rebuilding
  or deploying. Code retains the user's modified `AGENTS.md`; docs retain the
  user's modified Architecture v4.11 file and untracked `invSys dev notes.txt`.

# Decisions and constraints

- **Acceptable Inventory For Run** is the selectable external-stock allocation
  surface. `ReusableRunPaletteRows` intentionally omits a requirement with an
  incoming Recipe connection; do not make a routed intermediate look like an
  independently selectable managed-stock alternative.
- A routed intermediate is valid only after its upstream Process completes in
  the same active run. It is created under a fresh output `System_Key` and the
  downstream consume event later uses that exact key.
- Separate previously completed runs or general inventory do not satisfy a
  connected Recipe edge. Using existing general inventory instead would require
  an unconnected external requirement plus Ingredients Assignment and is not the
  intended full multi-Process acceptance path.
- The current UI has a real visibility gap: `ReusableRunManagerCheckRows`
  enumerates external `mAllocations` only, even though `BuildNodeConsumeItems`
  adds routed connection keys to the authoritative `USED` event.
- A single numeric **Used Goods** value must not sum unlike UOMs. Its mixed-UOM
  display requires an explicit contract decision before implementation.

# Evidence and traceability

- Visible symptom: the final Process showed only its external packaging stock
  in the palette; its connected upstream input showed **WAITING UPSTREAM** in
  the complete plan.
- Screenshot inference: blank active output keys and zero Process totals indicate
  that the currently loaded run had not recorded upstream completion. Re-verify
  whether the earlier successful Process runs were performed in another Recipe,
  run, batch, or form session.
- Known service path: `RequirementReadinessStatus` requires the source node to
  be complete and its exact output key sufficiently available;
  `BuildNodeConsumeItems` queues that same key as `USED`.
- Governing contract: Architecture v4.11 D15 Production Run - List requires
  same-run exact-key intermediate creation/consumption, Process-scoped
  readiness, remaining co-product balance, and event visibility.
- Last automated evidence: Slice 4aw source `7/7`, focused packaged public
  action `1/1`, packaged XLAM `81/81`, compile `142/142`, Slice 0 tooling
  `62/62`, full Release 1 chain `30/30`, and live roles `47/47`.
- This turn made documentation/handoff changes only, so D13 runtime RED/GREEN
  does not apply. The next runtime change requires a focused packaged RED through
  the actual Check In/Complete Run handlers.

# Do Not Repeat

- Do not add the upstream connection to the palette as a normal acceptable-item
  choice; that would let the operator allocate arbitrary stock instead of the
  run-owned upstream output key.
- Do not infer a run-session defect from the screenshot alone. First reproduce
  the exact uninterrupted operator sequence and inspect public-handler state.
- Do not add LB and EA into one numeric Used Goods total.

# Assumptions to Re-verify

- Whether the three reported successful Process completions occurred under the
  same loaded Recipe version, run ID, and active batch as the final Process.
- Whether form refresh, Process filtering, Recipe reload, or the released
  Process/Recipe version change cleared or replaced the active run session.
- Whether Inventory Check should be extended or a separate read-only routed-
  input projection is clearer; Architecture requires visibility but not a
  selectable palette row.

# Open questions and blockers

- Full four-Process Production acceptance has not passed end to end.
- Routed intermediate exact-key/quantity visibility is missing from the current
  operator projection even though the completion service consumes it.
- Confirm the mixed-UOM Used Goods presentation before implementing it.
- Remaining pre-release work: two-computer/two-warehouse Aggregator proof;
  comprehensive Viewer plus Admin detail/Action Path; SharePoint/GitHub/NAS
  cleanup and updater; and historical inventory worksheet import for analysis.
  The latter three require normative contract reconciliation.

# Immediate next action

Write a focused packaged test through the Production public form handlers that
completes an upstream Process, refreshes/selects its downstream Process without
reloading the Recipe, and records RED until the routed exact key is visibly
checked and then consumed.

# Critical references

- `src/Production/Modules/modProductionReusableRun.bas`
  - `ReusableRunPaletteRows`
  - `RequirementReadinessStatus`
  - `ReusableRunManagerCheckRows`
  - `BuildNodeConsumeItems`
  - `ProcessUsedGoodsQty`
- `src/Production/Forms/frmProduction.frm`
  - `RefreshReusableRunControls`
  - `mBtnManagerCheckIn_Click`
  - `mBtnManagerApplyOutput_Click`
- `tests/tooling/Test-Plan022Slice4avProcessScopedRun.ps1`
- Plan 022 Slices 4av, 4aw, and 4ax
- `invSys-Controls-v1.md` v1.39
