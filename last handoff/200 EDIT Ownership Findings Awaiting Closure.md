# EDIT ownership findings awaiting closure

## 1. Goal and release outcome

Complete4be-A, including Receiving B0, under Architecture4.11/Plan022/D13 without
regressing accepted Release1 behavior; broader4be-B remains separate.

## 2. Current verified state

Last verified2026-10-20T15:31Z. Code main31c27d40 and isolated branch
`work/d8-inv-edit-receipts` fa52aef8 are pushed. Docs main42a8b73 precedes this
closing commit. Controls1.691. [Draft PR4](https://github.com/soetrain/invSys_fork/pull/4)
remains unqualified. Main product source is fc2fba93; draft source is unchanged
from5d15970c. This turn changes evidence only; no implementation fix or deployment.
Preserve unrelated code AGENTS.md and docs067/deleted024/untracked023/024REPORT/025.

Focused `EditOwnership` on draft204d9be9 finished66 checks:60 PASS/6 FAIL, including
all five instrumented compiles. **Closure/restoration/native qualification is
pending.** Quit/COM release returned15:20:22Z, garbage collection15:20:22.982Z.
Windows reports Excel26956 exited but still enumerates it. Parent15848
(exec87688) and worker19576 remain live; the parent retains the original settings
snapshot. Do not terminate them, reboot, start Excel, or modify the isolated
worktree's pinned source/tooling before restoration and audit finish.
No process intervention occurred in this run.

Fresh cursor proof reached the exact target and restored its position. Passive
observer27104/exec59172 stopped normally at15:29Z, exit0; all failure counts were
zero. Its report is now stale. An interim native query found zero Excel events,
but does not replace the final audit.

## 3. Decisions and constraints

D8-INV-EDIT-01 remains approved, including all four tightenings; no approval is
pending. Preserve the single confirmed catalog/receipt/planned-event save before
queueing, exact event proof, reciprocal row binding and Inventory-private storage.
Modal Edit is unchanged. Preserve captured authority, System_Key, custom columns,
headless ownership and prior GREENs. No new ControlId/replay adapter.
Main workflow is authorized. Exit for error5, mouse lock or an approval request.

## 4. Evidence and traceability

Five direct failures: mixed ADD/EDIT fails whole-selection refusal and open-owner
preservation; pending quantity retry reaches queueing, changes source and fails
owner preservation. Neither failed action reports IMPORTED. Source inspection
locates the gap: Domain refuses borrowed catalog writes too late, while pending
quantity bypasses that Domain gate. Protect pending EDIT ownership in Core
preflight before any selected row writes, and recheck at write boundaries.

Passing: mixed retry with source closed creates one entity and one adjustment;
reciprocal receipt saved; completed imported inspection leaves source open and
unchanged; injected queue failure proves one catalog application and the receipt/
event plan saved before queueing. Final retry preserves catalog history and proves
exactly the original applied event/log effect.

The sixth failure, `QueueFailure.FreshRetryKeepsPlannedEvent`, expects a new
queue observation after the preceding defective borrowed retry already applied
the event. It is dependent, not proof of lost/changed identity; exact final
event/log proof passes. Keep this distinction when correcting the test or fix.
Eight Config boundaries preserve cells/shapes, Auth bytes and extraction.
Five preserve Config bytes; three metadata changes have nine observed bootstrap
saves. Older phase/result/save/headless Config gaps remain open.

Old excluded combined controller7990302e43dd4710bb3ada422e38182b subsequently
restored settings and all pins at10:21:01.698Z; audit10:21:21.713Z records one
Excel event. It remains excluded. Separate qualified baseline REDs remain valid.
Static draft metrics remain Run9/unresolved45/duplicates194/oversized27;
full provenance and retained GREENs are in code Inventory evidence.

## 5. Do Not Repeat

- Do not call the current result qualified before closure/restoration/native audit.
- Do not kill an exited Excel process or its restoration controller.
- Worker DeployRoot must be repository-relative: use
  `../invSys_fork/deploy/validation-tester-recording-02`, not an absolute argument.
  First launch8898531c7481433fa15de7a6c56af245 failed before Excel, restored all
  pins/settings with zero faults, and is not behavioral RED.
- Do not use `Get-AdminInventoryEntryFacts -Persisted` while observing borrowed
  ownership: it closes the inspected disposable source.
- Do not blanket-refuse completed imported inspection or modify modal Edit.
- Do not infer a graphics/RDP/native-warning cause from these results.

## 6. Assumptions to Re-verify

Process identities/removal, parent survival, restoration and all pins, final native
audit, Git/PR head. A fresh bounded cursor proof and observer are required before
another UI test. Current observer stop file exists.

## 7. Open questions and blockers

Windows cleanup currently blocks qualification and new Excel work. Native shutdown
delay cause remains unresolved. After closure: ownership fix/GREEN, all receipt
and mutation cases, retained regressions, final packaged/static/layout/live-role
recording/guide conclusion, Receiving B0/full Release1-chain and visible acceptance.
4be-A is not complete.

## 8. Immediate next action

Inspect the existing controller's closure and native audit, then—only after
qualification—fix pending EDIT ownership using the actual Upload RED and rerun
`EditOwnership` GREEN with fresh desktop proof and native resource tracing.

## 9. Critical references

- Architecture4.11 D8-INV-01/D8-INV-EDIT-01/D13; Plan022; Controls1.691.
- Worktree `../invSys_edit_receipts`; controller
  `reports/runtime/inventory-imported-controller/a2f98d5ae51a4d918b31954cb6d4e110`.
- Worker `reports/runtime/slice4be-admin-inventory-entry/b5f074dc846640b5814231788fc8d4cd`:
  `red.json`, `inventory-shutdown.jsonl`, `guide-resources.jsonl`, Config observations.
- Desktop `reports/runtime/edit-ownership-desktop-20261020.json` and `.cursor.json`;
  stop file `reports/runtime/edit-ownership-desktop-20261020.stop.txt`.
- `Test-Slice4beInventoryImportedRows.ps1 -InventoryCase EditOwnership -Phase GREEN
  -TesterBaselineOverlayForTest -TraceResourcesForTest -DeployRoot
  ../invSys_fork/deploy/validation-tester-recording-02 -DesktopMonitorReport <fresh report>`.
- `modInventoryEditCommands.Prepare/Execute`, `modInventoryEditReceipts.SaveCatalog`,
  `Slice4beInventoryEditOwnership.ps1`, `Slice4beInventoryEditReceipts.ps1`.
- `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`.
