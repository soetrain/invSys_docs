# Slice 4be contract proposal awaiting approval

## Goal and release outcome

Achieve Release 1 acceptance while preserving accepted behavior; active Slice
4be needs an approved comprehensive Events/detail/control-usage contract before
packaged D13 implementation.

## Current verified state

- Last verified 2026-09-06: code `main` at `b108313`, clean; only the existing
  source-regression report changed after runtime-source baseline `0afe05f`.
  Documentation `main` started at `cb5c958`; proposal commit was `d891edc`.
  This update accompanies the readiness-evidence commit; resolve its exact hash
  with `git log -1` rather than assuming an earlier deployment hash.
- Architecture v4.11, Plan 022, and controls catalog v1.58 now contain a clearly
  proposed, inactive Slice 4be amendment. Runtime implementation has not changed.
- The pre-existing user edit to `067 Partial Goal Action Path NAS Contract.md`
  remains uncommitted and must be preserved. It requests both models for
  comparison and notes networking/deployment and historical-analysis deferrals.
- No Excel process was listed during read-only inspection. Recheck before any
  build/deploy; no operational workbook or installed XLAM was touched.
- The goal is active and incomplete. Release 1 acceptance is not claimed.

## Decisions and constraints

- Current D18 still requires curated, user-authored, versioned training records.
  Current Viewer wording still defers comprehensive history until after R1.
- The latest request describes revealing actual user-side controls. The proposal
  would replace curated Save/Import/Export with recorded control usage, introduce
  Admin-configurable Event Detail, and bring comprehensive Events into R1.
  This replacement is a proposal, not an approved decision.
- Proposed capture is Admin-enabled, off by default, bounded to 64 actions per
  submission, NAS-only, non-executable, and separate from canonical authority.
  Viewer Show Action Path is optional. Proposed history is 5,000 published
  durable records with 100-row pages and explicit source/coverage/freshness.
- The user's request requires explicit normative approval before implementation.
  Do not use the lower-precedence comparison note to invent a mixed contract.
- Preserve D12/D13/D14/D15, captured-workbook binding, exact `System_Key`, unknown
  columns, launcher reuse, and all accepted role regressions.

## Evidence and traceability

- Read current pointers/specification and inspected relevant plan, catalog,
  handoff, Viewer public action/form, Core event reader and publication sources.
- Symptom/request: comprehensive history and actual control revelation missing
  -> current ten-column snapshot projection has no profile/path handlers
  -> D18 and Viewer scope conflict with requested target
  -> approval, then proposed `tools/validate_slice4be_event_viewer.ps1` through
  actual packaged Viewer/Admin/role handlers, with meaningful RED before code.
- Documentation-only validation: inspect complete proposal diff, whitespace,
  pointer destinations, consistent pending status and bounded defaults. No
  runtime RED/GREEN, compile, layout, live-role, full-chain or UAT is claimed.
- Continuation readiness check against code `0afe05f`: Slice 4bc source
  assertions 5/5 PASS; Slice 4w source assertions 12/12 PASS and its sanitized
  report refreshed to the existing three-tab contract. Five deployment package
  names/sizes/SHA-256/package-set values match the local manifest; this is not
  source/build or runtime proof. The old Viewer runtime report still has two
  tabs and cannot establish a current baseline.
- Plan 022 now maps source/action boundaries and known gaps. The still-pending
  proposal was corrected to retain every contributing event line, distinguish
  historical unknown time zones, and declare an Admin-owned profile writer
  rather than extend the existing Core config-write/D5 discrepancy.

## Do Not Repeat

- Do not mark an amendment approved because it was committed or pushed.
- Do not infer actual clicked controls from an event type or current-state row.
- Do not introduce SharePoint storage for Action Paths or migrate business data.
- Do not overwrite or stage the user's modified handoff 067.

## Assumptions to Re-verify

- Installed package/feed state and networking were not verified this session.
- Current GREEN reports are historical until rerun against the candidate package.
- A different Action Path choice requires revising the three synchronized
  proposal sections before proceeding; it is not a routine implementation choice.

## Open questions and blockers

- Explicit approval is pending for the proposed D18 replacement and comprehensive
  Events scope; comparison with the former curated model remains unresolved.
- Admin audit rows have no stable event IDs and can be station-local. Resolve
  their publication/correlation contract before implementing that family; do
  not manufacture historical IDs or treat current state as durable history.
- Physical Aggregator acceptance and broader Release 1 gates remain open.

## Immediate next action

Obtain the user's decision on the linked amendment, synchronize its approval
state, then write and run the focused packaged Viewer/Admin behavioral RED
before implementing the first approved Slice 4be behavior.

## Critical references

- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`: D18 and following proposed
  Slice 4be amendment; Viewer scope; D12-D15 and D19.
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`:
  current Slice 4be contract review and execution sequence.
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`: pending Slice 4be controls.
- Code: `src/Operations/Modules/modInventoryViewer.bas`,
  `src/Operations/Forms/frmInventoryViewer.frm`,
  `src/Core/Modules/modInventoryViewerData.bas`,
  `src/Core/Modules/modWarehouseSync.bas`,
  `src/Admin/Forms/frmAdminSettings.frm`.
