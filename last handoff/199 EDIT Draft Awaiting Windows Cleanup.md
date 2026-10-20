# EDIT draft awaiting Windows cleanup

## 1. Goal and release outcome

Complete4be-A, including Receiving B0, under Architecture4.11/Plan022/D13 without
regressing accepted Release1 behavior;4be-B remains separate.

## 2. Current verified state

Last verified2026-10-20T09:42Z. Code main b7cfe159 and isolated
`work/d8-inv-edit-receipts` 204d9be9 are pushed. Docs main a979817 precedes this
closing commit. Controls1.690. Draft [code PR4](https://github.com/soetrain/invSys_fork/pull/4)
targets main. Main product source remains fc2fba93; receipt implementation is
isolated, uncompiled and unqualified. No XLAM promotion or acceptance occurred.
Preserve unrelated code AGENTS.md and docs067/deleted024/untracked023/024REPORT/025.

**Do not start another Excel test.** Original controller14172 (exec84828) remains
alive with its original settings snapshot in memory. Excel24144 was terminated
at09:07:42.137Z but Windows still enumerates it (`HasExited=True`,140262 handles).
Blocked worker17008 was stopped at09:32:37.268Z; normal-Quit helper16560 is gone.
The parent waits for Excel removal before restoration; closure/native-audit files
are absent. Preserve the parent; do not reboot or kill it while state is unrestored.
Do not change original source/tooling while its pins remain active.

Desktop observer4188 (exec93149) stopped normally, exit0, at09:41Z; all desktop
failure counts were zero. Fresh bounded movement proof passed at08:55Z with
restoration. Subsequent observations were passive, not fresh input proofs.

## 3. Decisions and constraints

D8-INV-EDIT-01 is approved with all four tightenings. Core coordinates; Domain
owns private receipts. One confirmed Inventory save holds catalog, receipt and
planned event ID before queueing. Retry uses the original event ID; exact
AppliedEvents/InventoryLog proof and reciprocal row/receipt binding are required.
Modal Edit remains unchanged. Preserve captured authority, exact System_Key,
unknown columns, primitive boundaries and accepted GREENs. No new ControlId or
replay adapter; no approval pending. Main workflow is authorized.
Exit for error5, mouse lock or a proposal needing approval.

## 4. Evidence and traceability

Separate qualified REDs remain valid: phase41 PASS/12 FAIL, result50/20,
canonical-save83/28 and corrected binding147/32. Binding includes43 privacy and
42 Config/Auth/read checks. Nine receipt-side mutation cases remain unexercised
until the real command produces receipts. Full details and run IDs are in the
code Inventory evidence document.

The later combined baseline is **excluded**, not a completed gate: Excel Event1001
`RADAR_PRE_LEAK_64` at07:48:10.752Z falls inside its native audit interval.
Exact process attribution and cause are unestablished. After over45 minutes
without a new assertion, normal Quit could not obtain a COM response and the
test-owned Excel process was terminated under the user's standing authorization.
Package hashes were independently preserved beforehand. Interventions and missing
restoration remain explicit; do not promote partial assertions or weaken the audit.
The expanded phase/result/save Config comparisons remain unqualified.

Draft source5d15970c static scan:385 components/6710 procedures; literal Run9,
unresolved45, duplicate groups194, oversized modules27. Metrics do not regress;
all new components are at most231 lines. Changed PowerShell parses and diff checks
pass. This is not VBA compile, behavioral GREEN or packaged acceptance.

Next protecting test is actual Upload `EditOwnership`: mixed ADD/EDIT selection
must refuse before any write when canonical Inventory is borrowed; saved-catalog
quantity retry must also refuse; completed imported receipt inspection remains
allowed. Draft Domain refusal currently occurs too late for the mixed case, and
quantity retry can bypass it. Run this RED before changing ownership behavior.
`-TraceResourcesForTest` reuses observational native resource tracing and its
existing diagnostic bounds; it is not a fix for the earlier native warning.

## 5. Do Not Repeat

- Do not restart Excel, stop the restoration parent or repeatedly terminate the
  already-terminated Excel process. `HasExited` alone is not completed cleanup.
- Do not infer a native leak cause, exact stalled procedure or operator dialog.
- Do not use `Get-AdminInventoryEntryFacts -Persisted` while proving borrowed
  workbook preservation: that option deliberately closes the inspected workbook.
- Do not add blanket refusal of completed imported read-only inspection or change
  modal Edit. Do not broaden the ownership fix into unrelated reparse-path rules.
- Do not count compilation/harness failures as behavioral RED.

## 6. Assumptions to Re-verify

Process removal, parent survival, restoration and all pins; native audit must retain
the excluded run's failure. The stopped desktop observer is stale. Before a fresh
test, prove bounded cursor movement/restoration and start a new observer.
Check Git and draft PR head. Do not alter pinned files in a running controller's
worktree.

## 7. Open questions and blockers

Immediate blocker: Windows has not removed terminated Excel, so original settings
restoration cannot finish. The snapshot is currently only in the controller.
Native warning cause is unresolved. After cleanup: ownership RED/fix, draft compile,
all receipt GREEN/mutation tests, Config preservation, retained regressions, final
static maintenance, packaged/layout/live-role recording and guide conclusion,
Receiving B0/full Release1 chain and visible acceptance remain open.

## 8. Immediate next action

Inspect the existing controller for Excel removal and restoration/audit; only after
those finish, run the isolated `EditOwnership` RED with resource tracing and fresh
desktop proof.

## 9. Critical references

- Architecture4.11 D8-INV-01/D8-INV-EDIT-01/D13; Plan022; Controls1.690.
- Original controller:
  `reports/runtime/inventory-imported-controller/7990302e43dd4710bb3ada422e38182b`.
  Inspect `closure.json`/`native-audit.json` when created; local intervention
  records and `desktop-segment-2.json` preserve the excluded-run facts.
- Worker:
  `reports/runtime/slice4be-admin-inventory-entry/49f06987bd2d4f7586b372790131484f`.
- Isolated sibling worktree `../invSys_edit_receipts`, branch
  `work/d8-inv-edit-receipts`; original frozen packages
  `../invSys_fork/deploy/validation-tester-recording-02`.
- From isolated worktree, after cleanup:
  `Test-Slice4beInventoryImportedRows.ps1 -DeployRoot <absolute frozen package root> -InventoryCase EditOwnership -Phase RED -TesterBaselineOverlayForTest -TraceResourcesForTest -DesktopMonitorReport <fresh absolute report>`.
- `Slice4beInventoryEditOwnership.ps1`, `Slice4beInventoryEditReceipts.ps1`,
  `modInventoryEditCommands`, `modInventoryEditPlan`, `modInventoryEditProof`,
  `modInventoryEditReceipts`, `modAdminInventoryWorksheet`.
- `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`;
  isolated static report `reports/runtime/edit-receipts-b19676af-static`.
