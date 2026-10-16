# Approved Aggregation RED paused on desktop lock

## 1. Goal and release outcome

Complete Plan022 Slice4be-A recording/guidance and required early Receiving B0;
implement approved D8-AGG-01 test-first while preserving qualified Release1 work.
Broader4be-B/R1 acceptance remains separate.

## 2. Current verified state

Last verified2026-10-14 Pacific: code branch `codex/aggregation-authority`, commit
`bbe8c064`; docs branch `codex/tighten-d8-agg-01`, parent `0977c48` before this
closing documentation commit. Both use the `justinwj` GitHub fork remote;
soetrain docs writes were denied403. Controls1.585; catalog42/179 unchanged.
Product VBA remains unchanged from `ff5db426`. Isolated unpromoted baseline
`deploy/validation-aggregation-baseline-01` contains five source-built packages;
the older tester-recording-02 candidate is absent locally. No current/NAS deploy.

Goal is **paused under the user's explicit conditional stop instruction**. At
21:32:12Z (14:32:12 Pacific), owned monitor
`resume-d8-agg-01/write-boundaries-entry-12` records150 samples, one desktop lock,
zero error5/mouse-lock/other failures, and stops. The macro guard dispatches no
further work. Excel closes normally; settings/package/tooling pins are preserved;
delayed native audit is clean through21:32:41.5673421Z. Close Excel before builds.
Own tooling/evidence/docs changes are committed and pushed at session close;
runtime reports/packages are ignored. No unrelated dirty files were present at
the closing status check.

## 3. Decisions and constraints

The user explicitly approved tightened D8-AGG-01 on2026-10-14 Pacific and cleared
Handoff180's approval blocker. Approval is recorded in v4.11/Plan022/Controls;
do not request it again. D8-TESTER-01 remains approved/qualified.

D8-AGG-01 requires original workbook/live source context, Core-owned interactive
authorization against the captured source warehouse, joint source/authority/feed
checks at each named boundary, complete staging before promotion, and durable
honest retry evidence until verified explicit resolution. Changed saved trusted
feed requires reopening. Borrowed/read-only/overlapping output refuses; existing
scheduled headless authorization stays separate. Fixed feedback excludes entered
paths/values; D17 descriptors and exact System_Key identity remain. No Aggregation
recording ID or replay permission was added. Stop/return again on desktop error5,
lockscreen, mouse lock or a new proposal requiring approval.

## 4. Evidence and traceability

The code Aggregation receipt is authoritative for scope and runtime identifiers.
Completed unchanged-product RED:74 PASS/10 FAIL for packaged drive/UNC path-only
comparison;88/35 for post-inspection/output/retry boundaries;104/46 for broader
authority/refusal boundaries, all without harness exceptions and with preserved
closure. The104/46 gate has a documented decoy test-order caveat.

Normalization alone fixes the drive fixture's SaveAs1004; both UNC variants
persist exact entities/quantities, but original dispatch is noncanonical. Do not
generalize doubled-separator failure to UNC/NAS. Earlier65/7 on tester-recording-02
proved stale dispatch, not persisted unauthorized callback output.

Subsequent runs are **partial**, never completed RED gates:

- controller `284090309d21490c9035d371ffcb1089`, worker
  `5b3ca6761e7f4e1a9df54e9cedbde1db`:100/56 includes stale-monitor failure before
  NewSession completes; fresh decoy restores workbook-switch mutation RED.
- controller `9e24d501299449848c058b5acdcd0fb4`, worker
  `bf14b6b9fd8449658cec401ce9c5858b`:24/19 includes duplicate-fixture COM failure.
- controller `9e1b652822644f26b9b499dad0d2a61b`, worker
  `3fff33580ed6494db1a2294c3d2f7a5e`:19/11 includes the required desktop-lock stop.

Reached cases: source/feed/session/rights changes after inspection still write;
pre-save source/feed/sign-out changes still write; malformed identity schema can
report completion; borrowed output is closed; before-save failure deletes the
old feed, retains no destination/source/phase-bound receipt, and allows same-form
retry. Feedback exposes entered paths. These map to approved D8-AGG-01/D17/D14
and the actual Ribbon/form protecting tests. No exact persisted-output claim is
added for injected post-save failure. The direct Core baseline adapter is expressly
supplemental; it does not prove a new interactive API's authorization.

All five instrumented packages compile in the latest stopped run. Four changed
PowerShell files parse and closing whitespace/link checks pass. Partial controllers
preserve pins/settings and have clean delayed Excel native audits. No new product
GREEN, layout/static/regression acceptance, recording or4be-A closure is claimed.

## 5. Do Not Repeat

- Do not treat a harness/lock stop, compile failure or missing fixture as behavioral RED.
- Do not treat direct Core success as packaged-action success or a receipt as authority.
- Do not dispatch Application.Run directly to a UserForm method; the test now uses
  a standard-module typed bridge for feedback observation.
- The duplicate fixture now copies the generated ListRow with a destination;
  mixed numeric/string COM Value2 assignments failed before reaching the owner.
- Use native Windows Git for LFS-aware status/commit; Linux Git misreports downloaded
  operational fixture binaries as modified. Preserve those files.
- Do not restart automation, desktop helpers or bindings while the stop condition holds.

## 6. Assumptions to Re-verify

Desktop input/lock state, Excel absence, baseline/package/tooling pins and both Git
statuses before resumption. Monitor publication failed in an earlier attempt;
bounded sharing-error retries and numeric failure capture exist only in ignored
orchestration. Short clean monitoring does not prove long-run reliability.
Two-PC/two-NAS and scheduled headless behavior have not been retested here.

## 7. Open questions and blockers

Desktop lock triggered the user's stop rule; approval is **not** a blocker.
The focused write suite remains incomplete. Corrected duplicate-source,
unknown-column/overlap and remaining phase cases have not yet reached verification.
Interrupted/reopened retry requires a protecting test; an add-in unload/reload test
was only drafted outside the repositories and is not implemented/qualified.
New direct interactive Core API denial proof, product implementation/GREEN and
broader4be-A/Admin/B0 acceptance remain open.

## 8. Immediate next action

After the user resumes and desktop input is reverified, run the existing focused
write RED on the unchanged isolated baseline to completion, then protect restarted
retry and remaining applicable phases before editing product VBA.

## 9. Critical references

- v4.11 D8-AGG-01, D17, D14, D18-REPLAY-01; Plan022; Controls1.585.
- Code `tests/integration/plan022_slice4be_admin_aggregation_results.md` and
  `plan022_slice4be_admin_remaining_coverage.md` / `plan022_slice4be_remaining_acceptance.md`.
- `Test-Slice4beAdminAggregation.ps1 -DeployRoot deploy/validation-aggregation-baseline-01
  -Phase RED -WriteBoundariesOnly -DesktopMonitorReport <fresh-report>`;
  full matrix uses `-PathComparison -BoundaryMatrix`.
- `Slice4beAdminAggregation.ps1`, `Slice4beAdminAggregationBoundaries.ps1`,
  `Test-Slice4beConfigCommands.ps1 -AdminAggregationOnly`.
- `frmAggregationSources.mBtnAggregate_Click`,
  `modAdminConsole.RunHQAggregationFromSourceSet`,
  `modHqAggregator.GenerateGlobalSnapshotFromFiles`, `WriteGlobalSnapshotWorkbook`.
