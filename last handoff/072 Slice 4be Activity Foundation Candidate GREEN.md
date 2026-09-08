# Slice 4be activity foundation candidate GREEN

## 1. Goal and release outcome

Achieve Release 1 user acceptance under Architecture v4.11 and Plan 022,
including comprehensive Operations/Admin Events and How-To/Diagnostic/Compare
Action Paths while preserving every accepted workflow and architectural invariant.

## 2. Current verified state

- Last verified 2026-09-07: code main `03f7f20`, pushed and clean; documentation
  main `8cede80`, pushed. This handoff/pointer is committed afterward.
- Active slice is **4be.1 Coverage and activity foundation**, still incomplete.
  Only the first two control integrations are implemented: Admin Save Value and
  Production Retrieve UOM Catalog. The broader catalog, publication and later
  Settings/Viewer/recording/guide work remain required.
- Candidate packages: `deploy/validation-activity`. `deploy/current` is unchanged;
  no operational workbook or NAS rollout was changed. Candidate hashes are in
  the linked evidence. Rebuild at the final destination before promotion;
  copying XLAMs can retain candidate absolute references.
- No Excel process remained after final verification. Recheck before building.
- Preserve the user's unstaged handoff 067 (3 additions / 3 deletions) and
  untracked `expert guidance docs/023 Slice 4be Critique.md`. Neither was staged.

## 3. Decisions and constraints

Approved D18 remains active; no repeat approval is needed for this synthesis.
One shared observation foundation supports authored How-To, immutable Diagnostic
evidence and Compare both. No replay, repair, override or hidden command dispatch.
Semantic inheritance permits clarifications/new controls/stronger tests, not
architectural contradictions. Plan 022 remains current; critique 023 is advisory.

Core owns activity storage/read/policy boundaries and configuration commands;
workflow owners determine effects. The role controllers call the actual owners
and report fixed sanitized outcomes. Captured context includes warehouse/station,
runtime/config paths and invSys session version, but paths remain local opaque
state and never enter records. Re-authenticating the same actor invalidates the
old form command binding. Store failure alone must not block an authorized action.

The approved policy tables are `tblEventTrackingPolicies` and
`tblEventTrackingControls`; both absent means version 0 defaults. Partial or
invalid policy fails closed for optional tracking. JSON schema/catalog/policy/
ordinal fields are integers; UTC must be calendar-valid. Outside recording,
ordinal is zero. SHA-256 covers the complete body before its final hash property.
Core/Domain stay headless; exact System_Key and unknown columns remain protected.

## 4. Evidence and traceability

See [foundation evidence](../../invSys_fork/tests/integration/plan022_slice4be_activity_foundation_results.md)
for hashes, commands and limitations.

- Initial activity RED: 19 PASS / 12 FAIL, preserving all 18 D5 checks.
- Boundary RED: 62 PASS / 5 FAIL, exposing coercible JSON schema values,
  impossible record/policy UTC, invalid policy version and orphan ordinals.
- Real-handler stale-session RED: 68 PASS / 2 FAIL; both forms could still write
  after sign-out/sign-in. Role controller guards now enforce D18 before commands.
- Final focused **70/70 GREEN**, repeated after strengthening the corrupt-hash
  case to retain a valid 64-hex shape. Public handlers remain the protecting path;
  direct API tests supplement them.
- Candidate: **5/5 compile**, cold-start references PASS, **81/81 packaged**,
  **48/48 live-role**, **30/30 ordered full chain**, Viewer PASS, Production
  layout/window behavior PASS, **3/3 launchers**, **2/2 reusable Production**
  including fresh Excel restart and exact persisted Recipe/workbook reuse.
- Static regenerated: 1,077 candidates, 192 duplicate groups, 45 unresolved
  dynamic calls, eight literal Application.Run targets, 28 oversized ratchets.
  Auth shrank by two lines; Production form line count stayed 11,700.
- Actual Settings save/tracking-failure screenshot and all layout screenshots
  inspected. Human acceptance and the new Event Tracking/Action Path UI are not
  claimed. Runtime reports/screenshots remain ignored.

## 5. Do Not Repeat

- Do not treat these two controls as comprehensive 4be.1 coverage or release
  completion. All remaining 4be.1-4be.6 work stays in scope.
- PowerShell COM custom-property reads were unreliable; verify package metadata
  from XML. VBA reads the producing package properties correctly. The builder
  stamps closed staged files using its existing OpenXML dependency.
- Excel COM faulted once in live-role and once in full-chain; fresh serial runs
  passed without source changes. Confirm the specific handle is terminal, inspect
  recovered workbook ownership, and close only owned fixtures before retrying.
  A completed smoke validator left an empty Excel process; it was verified empty
  and closed. Never kill unidentified or operational Excel sessions.
- Minimum/default layout captures match because the form enforces its approved
  1110x800 minimum. A misleading preview was resolved by reopening the original;
  no screenshot-helper change was retained.
- `modAuthSession` is a new Core dependency. Explicit source-fixture imports were
  updated; keep packaged and source harness composition consistent.

## 6. Assumptions to re-verify

Git/Excel state, candidate hashes and runtime control reachability. Physical
NAS/multi-station UAT and human acceptance remain separate from disposable tests.

## 7. Open questions and blockers

No approval blocker. Full catalog/eligible handler coverage, source-event
correlation/publication, policy compatibility writes, Settings tabs/preferences,
comprehensive Viewer, recording/conclusions, guides and Compare both are unfinished.
Do not infer any of these from the initial foundation's passing tests.

## 8. Immediate next action

Create packaged RED through Receiving's existing real Confirm Writes handler seam
requiring activity attempts/results correlated to every emitted/applied source
EventID, while extending the maintained full control-coverage accounting.

## 9. Critical references

- Architecture v4.11 D18; Plan 022 4be.1-4be.6; controls catalog v1.62.
- `Core.modActivity`, `modActivityCatalog`, `modActivityPolicy`,
  `modActivityStore`, `modTrainingJson`, `modTrainingWire`, `modAuthSession`.
- `modAdminSettingsAction.SaveValue`, `modProductionUomAction.Retrieve`;
  actual Settings/Production handlers retain their operator entry points.
- `Test-Slice4beConfigCommands.ps1 -CheckActivityEvidence -CheckActivityFoundation`;
  `Slice4beActivityAssertions.ps1`, `Slice4beActivityFoundation.ps1`.
- Next seam: `frmReceiving.TestRunConfirmWritesActionForWorkbook` ->
  `mBtnConfirm_Click` -> `modReceivingPostingService.ExecuteConfirmWrites`.
