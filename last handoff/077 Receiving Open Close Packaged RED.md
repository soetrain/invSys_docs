# Receiving Open/Close packaged RED

## 1. Goal and release outcome

Achieve Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths,
while preserving accepted workflows. The full Goal remains active.

## 2. Current verified state

- Last verified 2026-09-07: code main **cb9c0dc**, docs main **595f3a8**, both
  pushed. This handoff/pointer is committed afterward.
- Active **4be.1 Coverage and activity foundation**, incomplete. Receiving
  Open/Close now has reproducible packaged RED: **531 PASS / 47 FAIL**, 578 total,
  preserving all **491** earlier GREEN checks with no harness exception.
- No runtime source, Ribbon generation or package changed. Catalog 4 candidate
  `deploy/validation-receiving-freshness` retains all five hashes and its previous
  complete checkpoint gates. Source runtime is still **130df29**. Excel is closed;
  recheck before builds or validation. Accepted deployment/NAS is not rolled out.
- Preserve unrelated docs: handoff 067 has 3 additions/3 deletions; critique 023
  remains untracked. Neither was edited or staged.

## 3. Decisions and constraints

D18 semantic inheritance and synthesized contract remain approved; no approval
blocker. Open/Close clarification was synchronized and pushed as docs **b22b3bd**
before implementation. Controls v1.72 and Plan 022 record current RED.

Catalog 5 is pending. RECEIVING_OPEN and RECEIVING_CLOSE belong to
RECEIVING_WORKFLOW and existing RECEIVE_POST eligibility. Actual Ribbon dispatch
observes REQUESTED and OPENED/REUSED/FAILED, with Unknown data effect and empty
references. Only the existing launcher resolves/provisions/binds the workbook.
Reuse requires valid captured session/warehouse and still-open workbook.

Explicit Close/window close each have one REQUESTED/CLOSED pair; CLOSED is
Info/Unchanged and never posts/clears staging. Internal replacement, programmatic
Unload, termination and workbook shutdown are not user clicks. Stale-context
dismissal remains possible without attributing it to a new session. Optional
tracking failure must be visible without retrying opening or blocking dismissal.
Disabled Ribbon actions retain their guard; comprehensive denial observation
coverage remains separately pending. D5/D12/D13/D14 and exact System_Key apply.

## 4. Evidence and traceability

[Lifecycle evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_lifecycle_results.md)
contains exact command, baseline reference, results and limitations.

- Actual generated `RibbonOnActionOperations` is invoked with a fixture class
  implementing Office.IRibbonControl. Unsaved access seams invoke actual
  Close/QueryClose handlers. Window-close testing proves handler semantics,
  not native title-bar/mouse dispatch.
- First completed RED: 515 PASS / 37 FAIL. Expanded RED: 531 PASS / 47 FAIL.
  Forty-two failures concern absent observation pairs; one proves stale-session
  form reuse; four prove missing optional-store-failure notices.
- Same-name workbook reopening already binds the correct live object on unchanged
  runtime. It is a regression guard, not a second confirmed binding defect.
- Staged keys/unknown values, source authority bytes, unrelated workbook,
  direct/internal exclusion, dismissal and older-policy compatibility pass.
  Three full-size lifecycle captures were inspected; human acceptance is pending.
- Static regenerated: 171 components, 5,465 procedures, 1,077 candidates, 192
  duplicate groups, 45 unresolved dynamic calls, eight literal targets. All 28
  size limits hold against prior HEAD; source scanner links new protecting tests.

## 5. Do Not Repeat

- An initial 75 PASS / 1 harness failure stopped at fixture credential rejection;
  it is not product RED. Root cause was not captured. Fixture credential fields
  now use text formatting and exact in-memory checks without outputting values.
- Do not infer zero workbooks from null COM properties or terminate unidentified
  processes. Cleanup completed and Excel exited by itself; no process was killed.
- Preserve all previous packaged/root/cleanup guards and cooperative Production
  observer cleanup described in handoff 076. Do not copy/relocate compiled XLAMs
  with absolute references; build a new candidate directory.

## 6. Assumptions to re-verify

Git/Excel state, package hashes and current handler ownership. Prior 076 records
the full 491-check candidate's gates and retained comparison deployments.

## 7. Open questions and blockers

No current approval/environment blocker. Catalog-5 runtime implementation and
GREEN/release gates remain open. Receiving navigation/selection and worksheet
reachability, all other Operations/Admin coverage, publication, Settings policy
and profiles, personal preferences, comprehensive Viewer, recording/conclusions,
guide management and both comparison presentations remain incomplete. Physical
multi-station/NAS UAT and human comparison remain required.

## 8. Immediate next action

Implement the recorded Open/Close RED through the existing launcher/form owners
and catalog 5, then build a new five-package candidate and run the 578-check
packaged suite to GREEN while preserving the 491-check baseline.

## 9. Critical references

- Architecture v4.11 D18 Receiving Open/Close; D5/D12/D13/D14; Plan 022 4be.1-4be.6;
  controls v1.72; handoff 076 for previous full release gates and harness cautions.
- `Test-Slice4beConfigCommands.ps1` with all activity/foundation/Receiving/staging/
  local/lifecycle switches; `Slice4beReceivingLifecycle.ps1`; linked evidence.
- `modTS_Received.ShowReceivingForm`, `IsReceivingLauncherFormReusable`,
  `ShowReceivingMessage`; `frmReceiving.mBtnClose_Click`, `UserForm_QueryClose`,
  `ResolveOperatorWorkbook`; `modActivityCatalog`, `modReceivingActivityCodes`.
- Ignored `reports/runtime/slice4be-receiving-activity/lifecycle-red.json`,
  `lifecycle-first-red.json`, `lifecycle-first-harness.json`,
  `coverage-lifecycle-open.png`, `coverage-lifecycle-reuse.png`,
  `coverage-lifecycle-sessionchanged.png`.
