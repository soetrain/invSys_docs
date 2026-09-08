# Receiving staging and disposition candidate GREEN

## 1. Goal and release outcome

Achieve Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths,
while preserving accepted workflows and authority/identity rules. Goal is active.

## 2. Current verified state

- Last verified 2026-09-07: code main **54ec2cb**, pushed and clean; docs main
  **0b0366a**, pushed. This handoff/pointer is committed afterward.
- Active **4be.1 Coverage and activity foundation**, incomplete. The Receiving
  Add Selected / Add Disposition / Confirm Dispositions checkpoint is GREEN,
  with **262/262** combined focused checks and all checkpoint gates passing.
- Candidate: `deploy/validation-activity`; five exact hashes are in the linked
  evidence. Accepted `deploy/current`, NAS deployment and operational workbooks
  are unchanged. No Excel process remains; recheck before builds/validation.
- Preserve user changes: handoff 067 has 3 additions/3 deletions; critique 023
  remains untracked. Neither was edited or staged.
- A clean detached comparison worktree remains at sibling
  `../invSys_staging_baseline`, commit **f813f2f**. Its pre-change runtime was
  built separately into code-repo `deploy/validation-staging-baseline`.

## 3. Decisions and constraints

D18 synthesis and semantic inheritance remain approved; no approval blocker.
The catalog-3 staging/disposition clarification was committed in Architecture
v4.11 before implementation (**2c9f354** in docs), with Plan 022 and controls.
Controls **v1.67** now records candidate implementation and checkpoint evidence.
Discovered controls/stronger tests inherit D18; contradictions or material
weakening still require an explicit approved architectural decision.

Six controls are registered: Settings Save Value, Production UOM Retrieve,
Receiving Confirm Writes, Receiving Add Selected, Add Disposition and Confirm
Dispositions. Catalogs 1/2 remain readable; their policies cannot implicitly
enable newly introduced controls. Comprehensive coverage is not implemented.

Add observations distinguish local STAGED/Changed, form REJECTED/Unchanged and
owner FAILED/Unknown. Staging has empty submission references. Confirmation
retains every exact source EventId and Submitted/Unknown state, with Unknown
Domain effect. Existing staging/posting owners execute once; observation cannot
retry, block or roll back authorized work. Stale session rejection remains
independent of optional tracking. Preserve captured workbooks, System_Key,
unknown columns, headless authority and D12's five packages.

## 4. Evidence and traceability

See [current candidate evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_staging_results.md)
for exact hashes, commands, outcomes and limitations.

- Pre-implementation actual-handler RED: **206 PASS / 15 FAIL**, expanded to
  **210 PASS / 35 FAIL**. All previous 190 checks remained GREEN.
- First GREEN **245/245**, expanded to **262/262** without further runtime edits:
  older policies, readable records, stale Add, store failure, exact source IDs,
  direct-service negative attribution and captured/unknown-column preservation.
- Five builds/explicit compiles/cold start; packaged **81/81**; live-role
  **48/48**; full chain/restart **30/30**; Viewer; Production three-size/five-page
  layout/window states; three launchers; reusable Production/restart all pass.
- Reusable Production was repeated in two independent full runs, each **2/2**,
  including new Excel processes. The revised-harness launchers pass **3/3**.
  Both full runs close Excel normally, preserve all five hashes and record no
  native crash event. No focused-only switches or weakened assertions.
- Eleven actual Receiving captures and three Production layout captures were
  inspected. These are automated visible evidence, not human acceptance.
- Static evidence regenerated: 1,077 candidates, 192 duplicate groups,
  45 unresolved dynamic calls and eight literal targets unchanged. All 28
  oversized-module ratchets hold; Receiving shrinks 1,259 -> 1,235 lines.

Earlier reusable Production runs crashed with Excel RPC failure/native
ntdll.dll c0000028 in both the candidate and independently compiled pre-change
runtime. The harness now signals its UI Automation dialog observer to stop,
waits for a terminal state, and rejects abnormal completion instead of forcibly
stopping the job. Subsequent full runs pass. This supports the harness hypothesis
but does not establish the native crash's cause; failures remain in the evidence.
Harness failures were never counted as behavioral D13 RED.

## 5. Do Not Repeat

- Do not discard the failed runs or infer that a pre-change failure alone proves
  the candidate safe. Keep cooperative observer shutdown and full assertions.
- Run Excel validators serially and wait for the existing process to finish.
  A null COM Application/Workbooks property is not proof of an empty workbook
  collection. Identify ownership/creation before handling a leftover process.
- Hash package files after Excel closes. Open fixture files require the existing
  shared-read hash helper; ordinary Get-FileHash can fail on Excel's file locks.
- Install unsaved fault seams before live sessions/forms; Core edits reset
  globals. Reapply isolated runtime/operator-root overrides afterward.
- Runtime reports/screenshots remain ignored; never commit raw fixture records,
  operational values, machine/user details or credentials.

## 6. Assumptions to re-verify

Git/Excel state, candidate hashes, handler reachability and isolated fixtures.
The native crash cause remains unproven; investigate further if it recurs.
Physical multi-station/NAS UAT and user comparison remain separate requirements.

## 7. Open questions and blockers

No approval/environment blocker. Receiving Refresh/Clear/Close/Open and optional
navigation/selection remain pending, as does comprehensive coverage across other
Operations/Admin surfaces. The hidden Receiving worksheet compatibility button
still needs reachability evidence or explicit caller coverage.

Publication, Event Tracking policy/detail Settings, personal view preferences,
comprehensive Viewer, recording/evaluation, guide authoring/search/version/
export/import and Compare both remain incomplete. Required canonical audits
and business owners remain effective throughout this work.

## 8. Immediate next action

Add and run packaged actual-handler RED for Receiving Refresh/Clear observations,
stale binding and direct/internal-call negative attribution, retaining 262 GREEN,
before extending the runtime catalog or changing those handlers.

## 9. Critical references

- Architecture v4.11 D18 staging/disposition and semantic-inheritance rules;
  D5/D12/D13/D14; Plan 022 4be.1-4be.6; controls v1.67 Receiving coverage map.
- `Test-Slice4beConfigCommands.ps1 -CheckActivityEvidence -CheckActivityFoundation
  -CheckReceivingActivity -CheckReceivingStagingActivity -CaptureEvidence`;
  `Slice4beReceivingActivity.ps1`, `Slice4beReceivingStaging.ps1`.
- `frmReceiving.AddSelectedInventory`, `modReceivingAddInput`,
  `modReceivingActivityAction.ConfirmWrites`, `modActivityCatalog`,
  `modReceivingActivityCodes`, `modActivityReferences`.
- Next: `frmReceiving.mBtnRefresh_Click` / `RefreshClicked` currently ignores
  the Boolean result from `modTS_Received.RefreshReceivingUiForWorkbook`;
  `mBtnClear_Click` calls `ClearReceivingFormStagingForWorkbook` and
  `modReceivingPostingService.ClearReceivingStaging`. A second-table Clear
  failure can follow a first-table change; do not infer clean rollback.
- `tools/validate_plan022_packaged_launchers.ps1`: `Invoke-PackagedCallback`,
  `Start-DialogCaptureAndDismiss`; ignored `slice4be-staging-production-*`
  reports retain failures, baseline hashes and both passing cooperative runs.
