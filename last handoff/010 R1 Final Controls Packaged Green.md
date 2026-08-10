# Goal and release outcome

Complete Plan 022 Slice 4 control acceptance under invSys v4.11: preserve the
working public Seed Demo Inventory workflow, correct the reviewed Viewer,
Boxing, Receiving-history, resize, and Shipping identity controls, and leave the
five packaged add-ins ready for operator acceptance.

# Current verified state

- Code: `main`, commit `1f7eacb` (`Complete R1 form control acceptance`), pushed.
- Documentation: `main`, based on `38b32e5`; this handoff, Plan 022, and
  `invSys-Controls-v1.md` are being committed/pushed at session close.
- Active slices: 4d Shipping visible `System_Key`, 4h Boxing usability/layout,
  and 4i Receiving Entries History are implementation/test GREEN.
- Excel is closed after validation. The five `deploy/current` XLAMs were rebuilt.
- The unrelated code `AGENTS.md` edit and normative-spec NAS-pointer edit were
  preserved and not included in the implementation/docs commits.
- Last verified: 2026-08-10.

# Decisions and constraints

- The Inventory Viewer is an Operations ribbon surface for signed-in users and
  uses a visible built-in table icon.
- Receiving history and item selection are separate: history reads `ReceivedLog`;
  a deduplicated selector supplies item metadata for new staging.
- Non-versioned inventory displays `NA`; only Shipping BOM relationships own
  `v1`, `v2`, etc.
- Physical inventory and its relationships use immutable string `System_Key`.
  Designs retain three-digit base-36 `DesignId`/`DesignVersion`; produced
  inventory still receives `System_Key`.
- Search/refresh/history are projections; local edit actions stage; explicit
  post/make/send actions queue events; only the processor writes inventory
  authority. See controls checklist item 4.

# Evidence and traceability

- Focused controls: 12/12 GREEN.
- Slice 10 Receiving: 10/10 GREEN.
- Slice 11 Shipping/Boxing: 11/11 GREEN.
- Shipping status anchor: 4/4 GREEN.
- Packaged Shipping layout/search/`NA`/identity: 1/1 GREEN.
- Packaged Receiving history durability: 1/1 GREEN.
- Packaged XLAM: 74/74 GREEN; RibbonX: 136/136 GREEN.
- Static baseline: 149 components, 4,585 procedures, 9 literal
  `Application.Run` targets, 48 unresolved dynamic calls, 185 duplicate-body
  candidates. The bounded `frmShipmentsTally`/`modTS_Shipments` growth exception
  is recorded in `tests/integration/plan022_slice4_controls_results.md`.

# Do Not Repeat

- Do not relabel numeric worksheet row values as `System_Key`.
- Do not use Receiving's history rows as inventory authority or as the item
  selector.
- Do not delete private legacy Shipping maintenance routines solely because the
  scanner reports them; reviewed reachability and protecting tests are required.

# Assumptions to Re-verify

- The operator must still visually verify Box Builder and Box Maker readability
  after grow/shrink/maximize/restore on the target display/DPI.
- Production's visible maximize/restore checkpoint remains pending even though
  packaged geometry is GREEN.
- The selected built-in Viewer icon should be visually confirmed in the user's
  installed Ribbon after Excel restart.

# Open questions and blockers

- No implementation blocker. Final Release 1 visual acceptance depends on the
  two manual checks above.
- Checklist item 4 wording is Codex's proposed control classification and may be
  adjusted by the user without changing the implemented authority boundary.

# Immediate next action

Install/restart the pushed packages, then visibly maximize/restore Production
and grow/shrink/maximize/restore both Boxing tabs while confirming the Viewer
ribbon icon.

# Critical references

- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
- `tests/integration/plan022_slice4_controls_results.md`
- `tests/tooling/Test-R1FinalControlAcceptance.ps1`
- `src/Shipping/Forms/frmShipmentsTally.frm`
- `src/Receiving/Forms/frmReceiving.frm`
- `tools/validate_plan022_packaged_launchers.ps1`
