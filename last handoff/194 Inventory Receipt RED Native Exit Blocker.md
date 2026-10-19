# Inventory receipt assertions reproduced; native exit blocked

## 1. Goal and release outcome

Complete4be-A under v4.11/Plan022, including early Receiving B0; Inventory
hardening still precedes final build and acceptance. Broader4be-B is separate.

## 2. Current verified state

Last verified2026-10-19 Pacific. Both branches main; code test commit a87aa09b,
docs preceding commit311b4ff. Closing evidence/handoff commits follow these.
Controls1.676. No product edits: src remains identical to e6b2809b/41ce313a.
No Goal mode was started. No XLAM build/save/deploy. No new approval pending.
Preserve unrelated code AGENTS.md and docs067/deleted024/untracked023/024REPORT/025.

**Cleanup remains live and unconfirmed.** At21:53Z, controller PID4676 (parent
10340), worker11344 and test Excel7740 still exist. Excel start is exactly
2026-10-19T21:33:38.9470807Z; reverify identity, never act on these PIDs blindly.
Controller retains the pre-test registry snapshot in memory. Do not start another
Excel test/build, discard the controller, or assume a reboot preserves that
snapshot. Its eventual closure/settings-restoration result must be inspected.
The separate passive desktop monitor was stopped intentionally; its stale
heartbeat cannot authorize another test.

## 3. Decisions and constraints

Continue normally, not Goal mode. The user permits opening/closing apps, including
Excel, but requires stopping on error5, mouse lock or a proposal needing approval.
Testing stopped after native TerminateProcess returned error5; this was not a
desktop-lock observation. Retain D8-INV-01/D8-AGG-01 approvals, captured source and
workbook authority, immutable exact System_Key, unknown columns, headless Core/
Domain, primitive envelopes and separate provisioned headless access.
No current Inventory build before focused GREEN. Source overlays are not packaged
acceptance. Main is the authorized local workflow; incoming Zenbook instructions
are machine-specific, not new remote defaults for this PC.

## 4. Evidence and traceability

Existing qualified imported-row RED195 remains46/10. New local controller
74af347e8cbf4338a11603bfcec49392 / worker evidence8890f62a4eb74c99b7a16bee96049de8
also executes all56 assertions:46 PASS/10 FAIL, five compiles pass, positive
prior-ADD control passes. Normal shutdown does not complete; therefore the local
run is unqualified. The worker-final result and native audit are not available.
Derived `assertion-log-red.json` preserves the exact assertion log explicitly.

Native exit: reported exit code0, .NET HasExited=True, but WaitForSingleObject258
and one running thread. Ordinary stop did not complete exit. Direct termination
at21:51:51.726Z (14:51:51 Pacific) returns FALSE/error5, handle still unsignaled.
No visible owned dialog. Cause unresolved; package-close order is only a hypothesis.
Desktop monitoring separately reports zero lock/error5 samples, with successful
bounded cursor movement before the run. No post-stop cursor test was attempted.

Independent hashes preserve5 packages/498 tooling/426 source files;498 tooling
parsers pass. Static baseline regenerated:377 components/6655 procedures/144014
lines; literal/unresolved Application.Run9/45; duplicate candidates194;27 oversized
modules. All six Config extractions are read-only, cells/shapes unchanged, Auth
bytes unchanged;18 observed Config saves/byte changes remain causally unresolved.
Settings restoration is separate and still pending. See the code Inventory
receipt's local continuation and Config diagnosis for exact cohort/source scope.

Upload symptom -> IMPORTED status blindly skips receipt proof -> D8-INV-01
items5/6 -> unchanged `Slice4beInventoryImportedRows.ps1` protects the next fix.

## 5. Do Not Repeat

Do not call the local assertion count a qualified run or infer shutdown from exit
code0/.NET HasExited. Do not rebuild packages, retry Excel or close unrelated
workbooks while cleanup is unresolved. Do not apply the historical catalog patch,
repeat resolved approvals, or waive Config differences as metadata.

## 6. Assumptions to Re-verify

Fresh process identities/handle signaling, controller/worker state and settings
restoration; Excel absence; native events; desktop/input access; both Git states.
The Zenbook identity baseline is absent here. The available frozen Tester baseline
requires the explicit34-source dependency overlay, not the old thirteen-source
default. Preserve these scopes separately.

## 7. Open questions and blockers

Native cleanup is the immediate local blocker. Upload receipt proof, EDIT retry,
primitive Admin/Core facades, Create authority/yields, Config diagnosis, final
regressions/static/packaged/recording/B0 and acceptance remain open. No product fix
was attempted in this continuation. No proposal needs approval.

## 8. Immediate next action

Read the owned controller/worker/Excel state and recover confirmed cleanup/settings
restoration before any new Excel automation, then resume the receipt-preflight fix.

## 9. Critical references

- v4.11 D8-INV-01/D13/D18-REPLAY-01; Plan022; Controls1.676.
- `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`.
- `tests/tooling/Test-Slice4beInventoryImportedRows.ps1`,
  `Slice4beInventorySourceOverlay.ps1`, `Slice4beRecordingLifecycle.ps1`.
- Runtime controller root above: `forced-test-excel-stop.json`,
  `native-owned-stop.json`, `desktop-stop-observation.json`, assertion log/pins.
- Worker root above: overlay scope, Config diagnosis and Auth comparisons.
- `modAdminInventoryWorksheet.PreflightInventoryWorksheetTable`,
  `modAdmin.ValidateWorksheetCreateReceipt`, `modInventoryCreateReceipts`.
