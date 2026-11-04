# 215 Encrypted Settings Recovery Before Windows Restart

## 1. Goal and release outcome

Complete4be-A under Architecture4.11/Plan022; recover from the blocked cancelled
Inventory test before its focused GREEN. Product acceptance remains incomplete.

## 2. Current verified state

Last verified2026-11-03 21:23Z. Excel remains in unfinished shutdown; the original
cleanup controller remains alive. Goal is blocked. Handoff214 contains product,
test, package and static evidence; no additional product test or build has run.

Code worktree `../invSys_edit_receipts`, branch `work/d8-inv-edit-receipts`,
607bb030 before this evidence update; docs `work/excel-shutdown-timing`,46c9f97
before this handoff. Publish updates to existing soetrain draft PRs
[code4](https://github.com/soetrain/invSys_fork/pull/4) and
[docs3](https://github.com/soetrain/invSys_docs/pull/3), both targeting main.
Preserve all unrelated changes listed in214.

## 3. Decisions and constraints

The user authorized recovering the controller's settings snapshot before further
cleanup. Its in-memory snapshot is now backed up with Windows CurrentUser DPAPI
under `%LOCALAPPDATA%/invSysTestRecovery/creation-cleanup-20261103-2102/settings.dpapi`.
Only the current Windows user and SYSTEM have file access. It is outside Git and
contains no plaintext file. **This supersedes214's loss-of-snapshot restart
blocker.** It does not establish that Windows restart has been requested, performed
or that settings have been restored. No reboot was issued in this recovery turn.

**User approval,21:26Z:** the user explicitly approved restarting Windows.
Independent recovery validation passes again before scheduling the restart.
Restart dispatch/result is recorded locally in the controller's
`restart-request.json`; verify actual boot state after reconnecting. Settings
restoration still requires the explicit post-exit step below.

Restore under the same Windows account after Excel exits. Do not delete the
encrypted copy or local recovery helpers before verified restoration. Do not
print/export plaintext settings. Product contracts and unassisted release-closure
requirements are unchanged; D13 behavioral RED is inapplicable to this recovery
bookkeeping. Original work-increment deadline remains22:14:05Z/2:14 PM Pacific.

## 4. Evidence and traceability

The PowerShell debugger attaches to the exact cleanup controller's running
cleanup loop. Seven original registry keys/seven values are recovered. Encryption,
restricted ACLs, exact decrypted bytes, registry value kinds and values all pass
round-trip comparison in memory. A separate process successfully validates
decryption, the invSys registry scope and original restoration-helper hash.
The debugger is detached and its helper console closed; the parent keeps waiting.
No registry writes occur during recovery. No secret values appear in reports.

## 5. Do Not Repeat

Do not confuse validation of the backup with restoration, or assisted cleanup with
release closure. Do not rerun product tests while Excel remains. Do not repeatedly
force the same terminating process, discard the backup, or reset the work timer.

## 6. Assumptions to Re-verify

Actual Excel exit, Windows restart state, same Windows user, encrypted backup/hash,
whether the parent already restored settings, and relevant source/package pins.
The desktop monitor remains stopped; a new cursor proof/monitor is required before
Excel automation. All product evidence retains its original candidate scope.

## 7. Open questions and blockers

Windows/Excel shutdown remains unresolved. A restart is now recoverable, but has
not been performed. Settings restoration, delayed native audit and focused
QuantityPhase140-check GREEN remain open, followed by all broader gates in214.

## 8. Immediate next action

After Excel exits or Windows restarts, validate and apply the recovered original
settings, verify restoration and package preservation, then resume214's focused test.

## 9. Critical references

Paths below are relative to main `../invSys_fork`:

- Validate only: `powershell.exe -NoProfile -NonInteractive -File reports/runtime/restore-cleanup-settings.ps1`.
- Restore only after Excel exits: same command with `-Apply`; it verifies
  backup identity, decryption, registry scope, original helper hash and readback.
- Controller `reports/runtime/inventory-edit-creation/green-7a879d5cf8a941ccb8695faed591c8ab`:
  `settings-recovery.json` (encrypted-copy pointer/hash and fixed checks),
  `cleanup-owner.json` (live process identities), `package-pins.json`;
  `recovered-settings-restoration.json` is produced only on successful explicit restoration.
- `tests/tooling/Slice4beRecordingLifecycle.ps1` owns the original restore function.
- Code audit `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`;
  handoff214 for exact RED/build/partial158 results and next QuantityPhase command.
