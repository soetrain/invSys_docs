# Slice 4be Receiving activity RED

## 1. Goal and release outcome

Achieve Release 1 user acceptance under Architecture v4.11 and Plan 022,
including comprehensive Operations/Admin Events and How-To/Diagnostic/Compare
Action Paths while preserving accepted workflows and identity/authority rules.

## 2. Current verified state

- Last verified 2026-09-07: code main `95e82e3`, pushed and clean; docs main
  `79b8ab5`, pushed. This handoff/pointer is committed afterward.
- Active **4be.1 Coverage and activity foundation**, incomplete. New Receiving
  focused RED is valid: **30 PASS / 8 FAIL**, zero harness exceptions.
- No runtime VBA, form implementation, build script or XLAM changed this turn.
  All five `deploy/validation-activity` hashes still match foundation `03f7f20`.
  Accepted `deploy/current` is unchanged; no NAS rollout or operational mutation.
- No Excel process remains. Check again before running/building.
- Preserve user changes: handoff 067 has 3 additions/3 deletions; critique 023
  remains untracked. Neither was staged or modified.

## 3. Decisions and constraints

D18 synthesis and semantic inheritance remain approved; no approval blocker.
Normative specification is unchanged by this test-only discovery. Controls
v1.63 registers pending `RECEIVING_CONFIRM_WRITES`, owner `RECEIVING_WORKFLOW`.
Confirmed submission and aggregate batch success cannot establish application
of every source event. Shared activity stores observations, never business work.

Core owns activity validation/store/read/policy; Receiving's actual owner must
supply exact event references and truthful outcomes. Capture session/workbook
binding independently of optional tracking, and keep tracking failure visible
without blocking otherwise authorized work. Preserve exact System_Key, unknown
columns, D5 command ownership and D12's five packages. Add and Confirm
Dispositions remain separate pending activity coverage.

## 4. Evidence and traceability

See [Receiving RED evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_activity_red_results.md).

`Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-activity -Phase RED
-CheckReceivingActivity` passes all 18 D5 checks and twelve Receiving guards.
Actual Add handlers create two durable keys/EventIds; the retained form invokes
real Confirm Writes while another saved workbook is active. Applied case proves
both events in inbox/applied/log with exact keys/quantities and cleared staging.
Pending case proves both queued, neither applied/logged, and staging retained.
Both preserve workbook binding, quiet UI, other-workbook bytes/sentinel content
and the unknown staging column header.

Each case fails four activity assertions: attempt/result, ActivityId correlation,
all exact source EventIds, truthful effect. These eight failures arise from
missing Receiving observations, not eight independent defects. PowerShell/diff/
new-link checks pass. No new implementation GREEN or broad gates claimed.

The previous foundation remains code `03f7f20`: 70 focused, five compiles,
81 packaged, 48 live-role, 30 chain, Viewer/layout, three launchers and two
reusable Production/restart aggregate checks passed. These are historical,
unchanged-package evidence, not a fresh comprehensive acceptance claim.

## 5. Do Not Repeat

- Install any unsaved test VBA instrumentation before creating live forms or
  capturing sessions. Editing Core afterward resets referenced-project globals
  and caused a test-only error 91 dialog. The final harness installs a gate
  early and toggles `TestReceivingActivityGate.SetPending` without code edits.
- Reapply fixture template/operator-root overrides after editing Core; project
  resets may discard them. Final harness does so before Generate Warehouse.
- Resolve the inbox through `modRoleEventWriter.ResolveInboxWorkbookPath`; Admin
  fixtures use their configured inbox directory, not an assumed runtime root.
- Hash open fixture files using read access with FileShare.ReadWrite. The test
  never saves an unrelated workbook to enable inspection.
- Excel changed the other workbook's Saved flag even though its bytes and
  sentinel content stayed unchanged. Separate those checks; do not infer a
  business mutation from the dirty flag alone or silently claim arbitrary
  workbook non-mutation from this bounded sentinel test.
- A failed test left a verified empty recovery Excel process; it was closed.
  Never kill an unidentified Excel process or restart a still-live validator.

## 6. Assumptions to re-verify

Git/Excel state, candidate package hashes, actual handler reachability and source
outcome semantics. Pending fault instrumentation is unsaved and test-only.
Physical multi-station/NAS UAT and human comparison remain separate requirements.

## 7. Open questions and blockers

No approval or environment blocker. Implement Receiving activity and all
remaining coverage, policy compatibility/editor, Viewer publication/detail,
recording/evaluation, guides and Compare both. The precise serialized source
reference validation must be documented under D18 before implementation; the
RED currently requires each reference's exact EventId. Do not infer terminal
application from the existing service's blanket batch-success report.

## 8. Immediate next action

Extend the RED with strict source-reference and stale-session/failure cases,
then implement the Receiving-owned result/source envelope and real Confirm
Writes observation boundary without moving business authority into Core activity.

## 9. Critical references

- Architecture v4.11 D18; Plan 022 4be.1-4be.6; controls v1.63.
- `tests/tooling/Slice4beReceivingActivity.ps1` and opt-in parent harness.
- `frmReceiving.mBtnConfirm_Click`, `TestRunConfirmWritesActionForWorkbook`,
  `SetOperatorWorkbook`, `UserForm_Initialize`.
- `modReceivingPostingService.ExecuteConfirmWrites`: owns validated state
  collection, queue, processor/refresh and staging cleanup; currently no
  observation/source envelope. `cReceivingWorkflowState.EventId/SystemKey`.
- `modRoleEventWriter.QueueReceiveEventBatchServer`: reports acceptedCount only
  after successful batch merge/save; failures may leave uncertain effects.
- `modOperationsPrimitiveBridge.RunBatchAndRefreshOperatorWorkbook` ->
  `modOperatorReadModel.RunBatchAndRefreshOperatorWorkbook`: aggregate result.
- `modActivity.BeginAction/FinishAction/MakeBody`, `modActivityCatalog`,
  `modActivityStore.ValidBody`: currently only two configuration controls and
  empty SourceEventRefs. `modProductionUomAction` demonstrates role orchestration.
- Foundation handoff 072 and its linked gate evidence for earlier details.
