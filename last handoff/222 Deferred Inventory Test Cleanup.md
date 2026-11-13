# Handoff 222: Deferred Inventory test cleanup

Date: 2026-11-06 Pacific. Predecessor: handoff 221.

## Goal and release outcome

Complete invSys Release 1 Slice 4be-A under Architecture v4.11 and Plan 022. This tooling increment allows development Inventory edit-creation tests to return a provisional result after Excel `Quit` while preserving the unassisted cleanup and final acceptance gates.

## Current verified state

Last verified: 2026-11-06 1:55 PM Pacific. Code worktree `invSys_edit_receipts` is clean on `work/d8-inv-edit-receipts` at `12517047e11e94ff333b331cc22d07ba25cf645d`, pushed to `soetrain/invSys_fork`. The main code checkout retains the user's unrelated modified `AGENTS.md`. Docs is on `work/perplexity-handoff218`, whose head before this handoff is `cf9d2e00301324d67603a82e9a006e374d549430`; its pre-existing deletion, edit and untracked files are untouched. Excel is closed and both Inventory desktop monitors stopped normally. No XLAM was promoted.

## Decisions and constraints

No Architecture, Plan 022, Controls, VBA or XLAM contract changed. `Invoke-InventoryEditCreationDeferred.ps1` is opt-in and presently covers the single Inventory edit-creation controller. `PROVISIONAL` and `FullAcceptance=false` never establish release GREEN. A pending or unsafe cleanup blocks another run; a failed test with confirmed Excel closure, settings restoration and package preservation may be followed by a new run. Sequential Excel tests still wait for the prior Excel process to close.

## Evidence and traceability

- Focused no-Excel D13: synchronous scaffold RED at 7.1 seconds versus the expected early return; expanded safe-failed-test gate RED; final **4/4 GREEN**. Four PowerShell files parse and `git diff --check` passes.
- Packaged `ApplyFailure` through the original local controller: **34/34 provisional PASS** at `Quit`, with owned Excel PID still live; an immediate second launch was refused. Original controller then exited 0 unassisted with settings restored and five package hashes preserved. Deferred run `9c53132ae2144d898645195d5ea2861d`.
- First tracked-controller dispatch failed before Excel creation because its deploy argument was absolute. Preserved failed run `02a3bd5afc074d82a30ccbff860983c4` records Excel closed, settings restored and packages preserved. Corrected tracked-controller run `a96ec7ab642c4ad197a5d2af1291022d` returned **34/34 provisional PASS**; final receipt followed about 197 seconds later with `TestPassed=true`, `CleanupComplete=true`, `FullAcceptance=false`, natural Excel closure, settings restored and packages preserved. Both live monitors recorded zero error-5, lock and mouse-failure samples. Exact details: `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`.
- Code commits: `b33e8de771c1c95ab0c1a1e1211d716375aeaee2` and `12517047e11e94ff333b331cc22d07ba25cf645d`.

## Do Not Repeat

Do not pass an absolute `DeployRoot` to `Test-Slice4beConfigCommands.ps1`; it resolves that argument relative to `RepoRoot`. Do not count a provisional result as final acceptance or start another Excel test during pending cleanup. The prior WPAExporter policy denial and unresolved shutdown root cause are unchanged. Automatic approval review also rejected deletion of one test-owned temporary folder without a stated rule; leave it untouched.

## Assumptions to Re-verify

Before another runtime test, verify Excel closure, desktop input, current package pins and a fresh desktop-monitor heartbeat. Code draft PR4 targets `soetrain/invSys_fork:main` and its head matched the new code branch when last checked. Docs draft PR5 still targets `work/excel-shutdown-timing` from the stale `justinwj` head; the published docs branch on `soetrain` does not update it.

## Open questions and blockers

The cause of intermittent 3–4 minute Excel native exit remains unresolved. The deferred lane is limited to the active Inventory edit-creation controller; other controllers still use synchronous cleanup. Broader Inventory regressions, live-role, visible-operator, full Release 1 chain and final static/acceptance gates remain open. Docs PR routing remains blocked by the prior automatic policy rejection.

## Immediate next action

After a fresh Excel/desktop/pin preflight, run packaged `ResultSaveFailure` through the tracked deferred Inventory controller, then verify its final natural-closure receipt before another Excel test.

## Critical references

Architecture v4.11 D13 and D8-INV-EDIT-01/02; Plan 022; Controls 1.731; `tests/tooling/Invoke-InventoryEditCreationDeferred.ps1`, `Run-InventoryEditCreationController.ps1`, `Test-InventoryDeferredCleanup.ps1`, `Test-Slice4beConfigCommands.ps1`; `deploy/validation-inventory-config-03`; code PR4.
