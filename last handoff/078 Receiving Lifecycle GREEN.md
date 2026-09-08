# Receiving lifecycle GREEN

## 1. Goal and release outcome

Achieve Release 1 user acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths,
while preserving accepted workflows. The full Goal remains active and incomplete.

## 2. Current verified state

- Last verified 2026-09-08: code main **61d401f**, docs main **93753f3**;
  this handoff/pointer is committed afterward. Push and verify both repositories
  before ending this session.
- Active **4be.1 Coverage and activity foundation**, incomplete. Receiving
  Open/Close technical checkpoint is GREEN: **596/596**, retaining every prior
  491-check and first lifecycle 578-check identity.
- Final five-package candidate: `deploy/validation-receiving-lifecycle-dismissal`.
  All package hashes remain unchanged after validation; accepted catalog-4
  `deploy/validation-receiving-freshness` hashes also remain unchanged.
  Accepted deployment/current and NAS were not rolled out.
- Excel process count is zero. Recheck before builds or validation.
- Code tree is clean after checkpoint commit. Preserve unrelated docs: handoff
  067 has 3 additions/3 deletions; critique 023 remains untracked. Neither was
  edited or staged. Only this handoff/pointer is newly authored afterward.

## 3. Decisions and constraints

D18 semantic inheritance and synthesized contract remain approved; no approval
blocker. Architecture, Plan 022 and controls **v1.76** are synchronized. The
Open/Close refinement was approved and pushed before implementation; committed
UI-dismissal clarification preserves the existing owner-fact rule.

Catalog 5 adds RECEIVING_OPEN/CLOSE under RECEIVING_WORKFLOW and RECEIVE_POST.
Actual Ribbon dispatch observes REQUESTED and OPENED/REUSED/FAILED with Unknown
data effect and empty business references. The existing launcher alone resolves,
provisions and binds; reuse requires the captured current context and exact live
workbook. Direct macros and internal lifecycle operations are not user clicks.

Explicit Close records dismissal after Unload. Native QueryClose synchronously
hides the form, records CLOSED (Info/Unchanged), and permits native teardown.
Launcher invalidation sets a flag without releasing the cached reference inside
the active QueryClose frame. No deferred completion or notice on a later launch.
Dismissal never posts or clears staging. D5/D12/D13/D14, immutable System_Key,
unknown columns, headless authority and captured-workbook binding remain intact.

## 4. Evidence and traceability

[Lifecycle evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_lifecycle_results.md)
contains exact commands, candidate hashes, RED/GREEN, failed runs and limitations.

- Original packaged RED: **531 PASS / 47 FAIL**; initial handler GREEN: 578/578.
  Native-window expansion exposed **592 PASS / 4 FAIL** before final 596/596.
- Final candidate: five VBE compiles/cold starts PASS; packaged smoke **86/86**;
  live roles **48/48**; full ordered Release 1 chain/reconciliation **30/30**;
  Viewer PASS; Production layout at three sizes/five pages and native window
  transitions PASS; Receiving and Production captures inspected.
- Revised native dialog observer: meaningful harness RED **6 PASS / 1 FAIL**
  (confirmation dismissed), then **7/7 GREEN**. NoEligible launcher **3/3** and
  complete ProductionReusable plus clean restart **2/2** pass on final candidate.
  No reduced Production flags; no Excel Application Error in that final run.
- Static: 171 components, 5,471 procedures, 1,077 candidates, 192 duplicate groups,
  45 unresolved dynamic calls, eight literal targets. All 28 existing size caps
  hold; Receiving form shrinks to 1,218 lines and modTS_Received stays at 1,575.
- Relevant source tests pass, including tooling 62/62, launcher 24/24 and
  Production restart 6/6. Human comparison/acceptance remains pending.

## 5. Do Not Repeat

- Cancelled native close plus nested Unload can report CLOSED while the form
  remains visible. Releasing its cached reference inside active QueryClose
  caused a VBE7 access violation; recording in Terminate deferred evidence until
  the next launch when the reference was retained. Neither approach is accepted.
- A vanished window alone is insufficient: verify Excel survives, dismissal
  records are immediate, reopen works and business authority stays unchanged.
- Two full Production runs failed with ntdll/RPC errors at different stages.
  Accepted catalog-4 comparison passed. Replacing broad UI Automation traversal
  with process-owned native modal observation produced the final complete GREEN;
  the ntdll root cause remains unproven. Preserve the failed runtime evidence.
- The observer may dismiss only a sole enabled OK button on the owned modal
  dialog; native OK may have ID 1 or 2. Never accept OK/Cancel confirmations or
  traverse ordinary form accessibility trees. Keep cooperative observer cleanup.
- Do not infer zero workbooks from null COM properties or kill unidentified
  Excel processes. Recovery files were retained. Do not relocate compiled XLAMs
  containing absolute dependencies; build each new candidate directory.
- Diagnostic lifecycle-only/unsaved mutation modes cannot replace full GREEN.
  Keep previous packaging/cleanup guards and stale-test cautions from handoff 076.

## 6. Assumptions to re-verify

Git/Excel state, candidate hashes, captured handler ownership and policy defaults.
Native observer GREEN does not establish the cause of earlier ntdll failures.
Visible captures are agent-inspected evidence, not human UAT approval.

## 7. Open questions and blockers

No current approval/environment blocker. Receiving optional navigation/selection,
worksheet-button reachability and launcher denial observation remain pending,
as do other Operations/Admin coverage and shared publication. Settings tracking
policy/detail profiles/preferences, comprehensive Viewer, recording/conclusions,
guide management, both presentations and comparison remain incomplete. Physical
multi-station/NAS/Aggregator UAT and human comparison are still required.

## 8. Immediate next action

Reconcile the approved optional Receiving navigation/selection controls with D18,
then create and run packaged actual-handler RED covering deliberate user actions,
programmatic selection/initialization exclusions, default-off policy and captured
context before implementing the next coverage increment.

## 9. Critical references

- Architecture v4.11 D18; Plan 022 4be.1-4be.6; controls v1.76; linked evidence.
- `modTS_Received.ShowReceivingForm`, `NotifyReceivingLauncherFormTerminating`;
  `frmReceiving.CanReuseFor`, `UserForm_QueryClose`, `mBtnClose_Click`;
  `modReceivingActivityAction.CloseForm`; `modActivityCatalog`.
- `Test-Slice4beConfigCommands.ps1` with all activity/foundation/Receiving/staging/
  local/lifecycle switches; `Slice4beReceivingLifecycle.ps1`;
  `Test-Plan022DialogObserver.ps1`; `tools/plan022-dialog-observer.ps1`.
- Ignored `reports/runtime/slice4be-receiving-activity/lifecycle-dismissal-green.json`;
  `reports/runtime/slice4be-lifecycle-native-observer-production/production-reusable-production.md`.
- Next controls: RECEIVING_PAGE_RECEIPTS/RETURNS/PURCHASING and selection IDs in
  the catalog. `mTabs_Change` guards mLoading; `mLstReceiveItems_Click` currently
  lacks an input-origin guard and refresh programmatically selects a sole item.
  Text/search/focus/scroll/resize/input values are explicitly excluded by D18.
