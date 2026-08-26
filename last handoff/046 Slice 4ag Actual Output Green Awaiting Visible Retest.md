# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4ag makes each reusable Production output's operator-entered actual quantity
authoritative for managed inventory and restores readable run-list identity.

# Current verified state

- Last verified 2026-08-26: code `main` at `c276586`; normative/plan/control
  docs at `bc1ce80`.
- Active slice: Plan 022 Slice 4ag, automated GREEN; visible Production Run -
  List retest is pending.
- Rebuilt five-package set is in `deploy/current`; existing startup registration
  points to that deployment; Excel is closed.
- Preserved uncommitted user work: code `AGENTS.md` and the NAS/server note at
  the top of Architecture v4.11.

# Decisions and constraints

- **Planned** is the released output definition after batch scaling; it is not
  the produced quantity.
- Every reusable output row requires its own positive **Actual Output** before
  Complete Run. The actual quantity creates that output's managed entity under
  its fresh `System_Key` and appears under **Last Actual**.
- A routed output's actual quantity cannot be lower than its committed
  downstream quantity. Excess remains managed intermediate/co-product balance.
- Next Batch clears new staged actuals but retains Last Actual history in the
  open run. Input consumption, output keys, ordering, scaling, and Domain/Core
  authority are unchanged.
- The operator's earlier `430` was ignored by the superseded build and that run
  persisted `632`; no existing inventory entity was silently rewritten.

# Evidence and traceability

- Focused D13 RED `0/6`; focused GREEN `6/6`.
- Packaged Production callback/restart `2/2` records
  `ActualOutputAccepted=True`, `LastActualDisplayed=True`,
  `ActualInventoryQty=True`, and `SystemKeyHeadersReadable=True` through the
  output-list click, textbox change, Check In, and Complete Run handlers.
- Regressions: XLAM `74/74`, Ribbon/compile `142/142`, live roles `47/47`,
  ordered Release 1 chain `30/30`, launcher contracts `24/24`, dedicated NAS
  `16/16`, deterministic static `19/19`, reviewed growth `13/13`.
- Static metrics: 153 components, 5,092 procedures, 1,038 candidates; literal
  and unresolved dynamic-call ratchets did not regress.

# Do Not Repeat

- Do not use Planned as created output quantity after an operator has supplied
  Actual Output.
- Do not restore the reusable-run behavior that ignores `txtOutputReal`.
- Do not call exact managed identity **Inventory ID** or render narrow clipped
  `System_Key` columns on Production Run - List.
- Do not retroactively change the earlier `632` entity without explicit user
  authorization and a supported correction path.

# Assumptions to Re-verify

- Confirm normal Excel startup loads the `deploy/current` hashes rather than a
  cached older add-in instance.
- Last Actual is retained across Next Batch in the active run; clean-restart
  reconstruction of historical Last Actual was not added to this slice.

# Open questions and blockers

- Visible confirmation of the corrected Actual Output, Last Actual, Planned,
  and readable `System_Key` columns is unresolved.
- Whether the old tea entity persisted at `632` should be corrected is
  unresolved and requires explicit user direction.

# Immediate next action

Open Production Run - List, start a new batch, enter an Actual Output different
from Planned, complete it, and confirm Last Actual plus the Viewer entity
quantity equal the entered actual.

# Critical references

- `src/Production/Forms/frmProduction.frm:StageSelectedReusableActualOutput`
- `src/Production/Forms/frmProduction.frm:CompleteProductionRun`
- `src/Production/Modules/modProductionReusableRun.bas:CompleteReusableRun`
- `src/Production/Modules/modProductionReusableRun.bas:ValidateReusableActualOutputs`
- `tests/tooling/Test-Plan022Slice4agProductionActualOutput.ps1`
- `tests/integration/plan022_slice4ag_production_actual_output_green_results.md`
- `tools/validate_plan022_packaged_launchers.ps1:ProductionReusable`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md` Slice 4ag
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` Slice 4ag and section 8.4
