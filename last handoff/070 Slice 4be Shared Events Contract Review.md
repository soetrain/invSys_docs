# Slice 4be shared Events contract review

## 1. Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11 and Plan 022,
preserving accepted behavior while defining comprehensive Operations/Admin
Events and Action Paths for How-To training and diagnostic conclusions.

## 2. Current verified state

- Last verified 2026-09-07: code main remains `b4ce6d9`, clean and synchronized;
  documentation main contains proposal commit `e85dd39`. This handoff/pointer
  is committed afterward; resolve its commit with Git.
- No runtime code, tests, XLAMs or operational workbooks changed this session.
  No Excel process was found during inspection; recheck before build/deploy.
- Preserve the user's unstaged 3-addition/3-deletion edit to
  `067 Partial Goal Action Path NAS Contract.md`. It was not edited or staged.
- Slice 4be is at detailed-contract approval. D5 is implemented; this new
  Event Viewer/Action Path proposal is neither implemented nor user accepted.

## 3. Decisions and constraints

- User explicitly requires Operations/Admin coverage, a dedicated Settings tab
  controlling tracking, How-To and diagnostic uses, both methods for comparison
  and a user choice in Settings.
- Asked to distinguish the historical pairs, the user directed a synthesis:
  "take the best from each and discard the rest". Do not demand selection of
  mutually exclusive legacy models again.
- Proposed synthesis: one tracking foundation and one versioned Action Path,
  with How-To, Diagnostic and Compare both views. Keep authored guide editing/
  search/version/export/import plus immutable observed controls and outcomes.
  These detailed presentations/defaults/limits remain pending approval.
- The 2026-09-07 amendment following D18 supersedes only the inactive
  2026-09-06 replacement-only proposal. Current normative D18, D19 and Viewer
  scope remain effective until approval. Do not silently activate the draft.
- D5 Core read/command separation stays approved; Admin owns policy/profile UI,
  Operations owns personal preference UI, and Core/Domain stay headless.
- Preserve exact System_Key, unknown columns, captured workbooks, launcher
  reuse, and all accepted Seed/Receiving/Shipping/Boxing/Viewer/inventory/
  reusable Production behavior. No operational runtime or NAS rollout is
  authorized solely by approving a product specification.

## 4. Evidence and traceability

- Documentation-only validation: 40 checks passed for unchanged effective
  D5/D12/D13/D18/D19 sections, synchronized draft terms/limits, pointers and D5
  evidence existence, and clean code. Diff/whitespace reviewed. This is not
  behavioral RED/GREEN and does not claim new runtime acceptance.
- Shared Admin-history gap: `modAdminConsole.AppendAuditEntry` writes station
  `tblAdminAudit` without stable event identity. Proposed new activity store
  addresses future eligible actions; old station history stays unavailable.
  Protect with 4be.1 packaged Admin action/publication tests after approval.
- `frmAdminSettings` currently uses one runtime canvas. Proposed General/Event
  Tracking tabs and Operations personal preference surface need 4be.2 tests.
- Prior [D5 evidence](../../invSys_fork/tests/integration/plan022_slice4be_d5_config_commands_results.md):
  18/18 focused, 81/81 packaged, 48/48 live-role, 30/30 full-chain, five explicit
  compiles plus Viewer/layout/launcher/reusable Production checks. These remain
  the baseline, not fresh results from this documentation session.

## 5. Do Not Repeat

- Do not reinstate the rejected replacement-only proposal or treat the older
  "pick one" language as higher authority than the user's synthesis direction.
- Do not treat a clicked control, accepted submission, current-state supplement,
  edited guide or imported example as proof of local Domain completion.
- Do not invent historical Admin IDs or UTC from existing local Now fields.
- Keep Excel validators serial; compile packages explicitly and check cold-start
  references before preloading Core. Do not copy XLAM candidates with retained
  absolute dependency references into deploy/current.

## 6. Assumptions to re-verify

Git/package/Excel state before runtime work; source-to-control coverage and
completion sources during the first focused test slice; physical NAS/station
acceptance remains separate from isolated fixture evidence.

## 7. Open questions and blockers

Detailed approval of the new amendment is pending, including default collection,
explicit recording, personal view preference, diagnostic expectation rules and
bounds. Fresh human UAT and overall Release 1 acceptance remain open. No goal
status was changed during this contract-drafting session.

## 8. Immediate next action

After explicit approval and activation of the synchronized contract, establish
the packaged baseline and run 4be.1's first public Operations/Admin handler test
to observe missing stable activity/correlation behavior before implementing it.

## 9. Critical references

- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`: proposed Slice 4be amendment
  after D18; D5/D12/D13/D19 and Viewer scope.
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`:
  current Slice 4be, proposed subslices 4be.1-4be.6 and acceptance cases.
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`: version 1.60 pending controls.
- Code `modInventoryViewer.OpenInventoryViewer`, `frmInventoryViewer`,
  `frmAdminSettings`, `modAdminConsole.AppendAuditEntry`, `modConfigCommands`.
