# Handoff 218: Shutdown isolation and Codex test request

Date: 2026-11-03 Pacific. Predecessor: handoff 217, followed by the merged code PR6 and docs PR4.

Excel shutdown is the immediate priority. Perplexity implemented and source-tested a bounded diagnostic experiment; Codex is the test executor only. The shutdown cause, Inventory GREEN, and Slice 4be-A acceptance remain unresolved.

## Implemented increment

- **Review PR:** [code PR7](https://github.com/soetrain/invSys_fork/pull/7), draft, targeting `work/excel-shutdown-timing`, not `main`.
- **Diagnostic commit:** `62b3da4dbf7e39e0cf626e44f8949f1952a7b03b`, parent `97716fa2d3ca69e5fbff63f6a8859235a5576927`.
- **Local diagnostic checkout:** `C:\Users\Justin\repos\invSys_shutdown_diagnostic`, branch `work/perplexity-shutdown-table-isolation`.
- **Current worker/helper commit:** `44557070b2f1c3d7d744e1f8040315a837933a8f`, in `invSys_edit_receipts`, still on `work/d8-inv-edit-receipts`.
- **Main fixture/deployment checkout:** `invSys_fork`, still `main` at `6d247543c557cca92fb6324f484b6ec2087e4527`; its unrelated AGENTS edit is preserved.
- **Source verification:** new suite 30/30 on PowerShell 7.6.9 and native Windows PowerShell 5.1; also 30/30 against the older worker on Windows. Existing native timing suite 14/14, using a synthetic PowerShell child rather than Excel.

The previous instrument refused the current worker before dispatch because its GC trace layout had changed. It now accepts exactly the reviewed combined layout or split layout while preserving the existing trace marker and all original cleanup calls. Unknown layouts still refuse. This was a tooling compatibility failure, not a new packaged product RED.

The opt-in `OwnedTableTraversal` variant changes only the generated worker's `Table()` helper. It replaces nested COM enumeration with indexed acquisition and releases owned intermediate references; it does not release the borrowed workbook or returned table. Baseline leaves that helper unchanged. No product source, XLAM, final Quit/GC sequence, security setting, or operational workbook was changed.

The retained unassisted diagnostic measured Quit at 5.8042 ms and finalizer wait at 334012.7779 ms. This locates the observed delay, not its cause. Retained worker/helper versions differ from today's source, so compare a fresh baseline with the variant, never the variant alone with historical timing.

## Authorization and time boundary

Perplexity authorizes exactly one sequential pair below, subject to all preconditions. This does not authorize broader runtime tests, implementation changes, retries, package rebuilding, merges, or assisted cleanup.

- **Execution lane:** normal interactive Windows session on X1-Pro-Ai. Do not use PerplexitySandbox or run the controller through a short-lived remote-command executor.
- **Latest pair start:** 2026-11-04 04:40 UTC, Tuesday November 3 at 8:40 PM Pacific. If missed, remain paused and request a refreshed execution window.
- **Pair budget:** 60 minutes from preflight start, including cleanup and audit. Treat 25 minutes for either controller as an overrun: do not start another run, report the live stage, and preserve the controller that holds the settings snapshot.
- **Dispatch cutoff:** no new variant after 30 minutes from preflight start. Reserve the remaining time for completion, restoration, hash checks, and delayed audit.
- **Increment deadline:** 2026-11-04 05:40 UTC, November 3 at 9:40 PM Pacific. This is the existing fixed 4h44m cap, not a restarted clock.
- **Cleanup caveat:** the existing controller waits for natural Excel exit without a kill timeout. The time limit is a testing/escalation boundary, not permission to kill Excel or close the parent console. If natural cleanup remains blocked, stop further testing, leave restoration state alive, report the blocker, and obtain explicit approval for any targeted assistance. Such a run is excluded from release qualification.

Keep one Excel controller running at a time. Its owned worker and passive desktop monitor are expected; no parallel Codex test, manual Excel session, package builder, or second controller is permitted. Freeze the pinned inputs until the entire pair and final preservation checks finish.

## Exact frozen inputs

The local manifest is `invSys_fork\reports\runtime\perplexity-shutdown-isolation-source\frozen-input-pins.json`. A companion copy is retained beside this handoff as `218 Shutdown Input Pins.json`.

The verified local file contains 535 SHA256 entries. Its exact-byte SHA256 is:

```text
F2D2768D8B9B430E2B0D823243E32F75C973E42778893A8C6832D6B16C174B73
```

Use that existing local manifest; do not silently regenerate pins or normalize files to make a mismatch pass. It covers all 521 worker-folder PowerShell files, six diagnostic/source-test files, two monitor inputs, five XLAMs, and the historical assertion report. A Git checkout or download may change line endings; any resulting mismatch must be returned for review, not worked around.

Current worker SHA256:

```text
5D2D2192F2DAE05F67310A42AF9FA039ABE5E05EFCBB6613EC3C251D066EA4F3
```

Candidate directory: `C:\Users\Justin\repos\invSys_fork\deploy\validation-inventory-uom-02`. Do not substitute `validation-inventory-config-02`.

| Package | SHA256 |
|---|---|
| invSys.Admin.xlam | `86E792B7E4082D0F7992191922C003C2AF63D08C93CE99D40D8582D2DE5CF776` |
| invSys.Core.xlam | `959D72ED95EEA60D62A863F2E3F851D71DD699C9483496B1B63F59667292B7A7` |
| invSys.Designs.Domain.xlam | `67AA48E24AA6C5C6C7DC33FBC49776CDC0235FA8A7A9311D1DEA9D9276047322` |
| invSys.Inventory.Domain.xlam | `11F9476BD09CC48D38EC9284B6CF910100CBFA7BDFD4C343D7D12E625E409EE4` |
| invSys.Operations.xlam | `F81BCE2DB6DDAE2ACB23D079493223A9C91B9C5A43074E012911D4258FADCFC0` |

## Codex execution request

### Read-only preflight

Run in one normal-user Windows PowerShell session. Preserve these variables for the remaining commands; do not change source branches or pull into the frozen checkouts.

```powershell
$ErrorActionPreference='Stop'
$requestStarted=[DateTimeOffset]::UtcNow
if($requestStarted -gt [DateTimeOffset]::Parse('2026-11-04T04:40:00Z')){
    throw 'Authorization window expired; runtime remains paused.'
}
$roots=@{
    Main='C:\Users\Justin\repos\invSys_fork'
    Worker='C:\Users\Justin\repos\invSys_edit_receipts'
    Diagnostic='C:\Users\Justin\repos\invSys_shutdown_diagnostic'
}
$pinsPath=Join-Path $roots.Main 'reports/runtime/perplexity-shutdown-isolation-source/frozen-input-pins.json'
$pinsHash='F2D2768D8B9B430E2B0D823243E32F75C973E42778893A8C6832D6B16C174B73'
function Assert-FrozenInputs {
    if((Get-FileHash -LiteralPath $pinsPath).Hash -cne $pinsHash){throw 'Pin manifest changed'}
    $pins=Get-Content -LiteralPath $pinsPath -Raw|ConvertFrom-Json
    $commits=@{Main=$pins.MainCommit;Worker=$pins.WorkerCommit;Diagnostic=$pins.DiagnosticCommit}
    foreach($key in $commits.Keys){
        $path=$roots[$key]
        $head=git -c "safe.directory=$path" -C $path rev-parse HEAD
        if($LASTEXITCODE -ne 0 -or $head -cne $commits[$key]){throw "Commit mismatch: $key"}
    }
    if($pins.Files.Count -ne 535){throw 'Pin count mismatch'}
    foreach($pin in $pins.Files){
        $path=Join-Path $roots[$pin.Root] $pin.RelativePath
        if((Get-FileHash -LiteralPath $path).Hash -cne $pin.SHA256){
            throw ('Frozen input changed: '+$pin.Root+'/'+$pin.RelativePath)
        }
    }
    $expected=@($pins.Files|Where-Object Root -eq 'Worker'|ForEach-Object RelativePath|Sort-Object)
    $actual=@(Get-ChildItem (Join-Path $roots.Worker 'tests/tooling') -Filter '*.ps1' -File|
        ForEach-Object {'tests/tooling/'+$_.Name}|Sort-Object)
    if(@(Compare-Object $expected $actual).Count){throw 'Worker helper membership changed'}
    foreach($key in @('Diagnostic','Worker')){
        $path=$roots[$key]
        $status=@(git -c "safe.directory=$path" -C $path status --porcelain)
        if($LASTEXITCODE -ne 0 -or $status.Count){throw "Frozen source checkout is not clean: $key"}
    }
    'FROZEN_INPUTS_OK 535'
}
Assert-FrozenInputs
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel; do not dispatch'}
$worker=Join-Path $roots.Worker 'tests/tooling/Test-Slice4beConfigCommands.ps1'
$controller=Join-Path $roots.Diagnostic 'tests/tooling/Invoke-ExcelShutdownDiagnostic.ps1'
$deploy=Join-Path $roots.Main 'deploy/validation-inventory-uom-02'
```

Confirm no other controller is active, no build/edit is scheduled, and the desktop is unlocked. The unrelated main AGENTS edit and unrelated docs changes are not failures and must be preserved.

### Start the passive monitor

The existing monitor makes one bounded cursor movement and restores its position, then performs passive desktop checks. Run it once for this pair, from the main checkout.

```powershell
Set-Location $roots.Main
$pairRoot=Join-Path $roots.Main ('reports/runtime/excel-shutdown-isolation/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $pairRoot|Out-Null
$monitorReport=Join-Path $pairRoot 'desktop.json'
$stopFile=Join-Path $pairRoot 'monitor.stop'
$monitorScript=Join-Path $roots.Main 'reports/runtime/monitor-inventory-desktop.ps1'
$monitor=Start-Process powershell.exe -PassThru -WorkingDirectory $roots.Main `
    -ArgumentList @('-NoProfile','-File',('"'+$monitorScript+'"'),
        '-Report',('"'+$monitorReport+'"'),'-StopFile',('"'+$stopFile+'"'),'-Hours','1') `
    -RedirectStandardOutput (Join-Path $pairRoot 'monitor.stdout.txt') `
    -RedirectStandardError (Join-Path $pairRoot 'monitor.stderr.txt')
```

Wait at most 20 seconds for the first report. Before each controller, require its timestamp to be no more than 15 seconds old, the monitor process still alive, and `Error5Samples`, `LockSamples`, and `MouseLockSamples` all zero. Any missing, stale, future-dated, or failing proof stops dispatch; do not alter security or restart the monitor to erase a failure.

### Baseline command

Re-run `Assert-FrozenInputs` and confirm Excel count zero immediately before this command. Record the printed diagnostic directory and controller exit code.

```powershell
powershell.exe -NoProfile -File $controller `
    -RepoRoot $roots.Main -WorkerScript $worker -DeployRoot $deploy `
    -DesktopMonitorReport $monitorReport -TableTraversal Baseline `
    2>&1 | Tee-Object (Join-Path $pairRoot 'baseline-controller.txt')
$baselineControllerExit=$LASTEXITCODE
```

The controller prints `Shutdown diagnostic: <directory>`. Its `worker.log` prints `Inventory entry evidence: <directory>`; the latter contains `red.json`. Use those exact paths, not whichever folder happens to be newest.

Before the variant, require complete natural closure, restored settings, unchanged pins, fresh desktop proof, and all four completed timing intervals. Compare the fresh baseline's ordered `Check`/`Passed` rows with the pinned historical report at `reports/runtime/slice4be-admin-inventory-entry/060f370b1fb64ecd97505c881812e96d/red.json`.

The historical expectation is 78 rows, 50 PASS and 28 known FAIL. All identities, order, and outcomes must match before proceeding. Worker exit code 1 from those known failures is expected; missing reports, new failures, changed outcomes, compile/harness exceptions, or controller failure are not. Any difference stops the pair for Perplexity review, even if a changed result appears better.

### Variant command

Run only after the baseline gate passes, less than 30 minutes have elapsed since `$requestStarted`, and at least 30 minutes remain before the increment deadline. Re-run `Assert-FrozenInputs`, confirm Excel count zero, and recheck the same live desktop monitor.

```powershell
powershell.exe -NoProfile -File $controller `
    -RepoRoot $roots.Main -WorkerScript $worker -DeployRoot $deploy `
    -DesktopMonitorReport $monitorReport -TableTraversal OwnedTableTraversal `
    2>&1 | Tee-Object (Join-Path $pairRoot 'variant-controller.txt')
$variantControllerExit=$LASTEXITCODE
```

Require the variant's complete ordered `Check`/`Passed` sequence to equal the fresh baseline, not just its totals. Require manifest mode `DiagnosticHelperLifetime`, selector `OwnedTableTraversal`, and the same original worker, helper, diagnostic, and package pins. Generated worker hashes should differ because the helper and per-run paths differ.

Do not promote the variant based on a shorter time alone. One clean pair is exploratory evidence; Perplexity will decide whether controlled replication or a different hypothesis is warranted.

### Cleanup and evidence return

Let each controller complete natural closure and restoration before starting anything else. Do not kill Excel, close the parent console, remove GC calls, dismiss unexpected dialogs, retry failed actions, change settings manually, or delete fixtures while the worker/controller remains alive.

After natural completion, re-run `Assert-FrozenInputs` and check Excel count zero. Signal the passive monitor to stop using its exact `$stopFile`, then wait for that monitor process to exit; do not terminate unrelated processes. Keep the final monitor report and cursor-baseline report.

Collect an immediate and a delayed Application-log audit for Excel events 1000, 1001, and 1002 covering the pair's start through at least two minutes after the final native exit. Perform hash and JSON comparisons during the delay. If event-log access fails, report the audit as unavailable, not as zero events. Keep raw event messages local; return sanitized time, provider, event ID, record ID, and counts.

For each diagnostic directory, retain:

- **Identity:** `manifest.json`, `owner.json`, `package-pins.json`, `tooling-pins.json`, `diagnostic-tooling-pins.json`, and generated-worker hash.
- **Timing:** `phases.jsonl`, `timings.json`, `native-exit.json`, controller/worker wall times, and the worker's inventory shutdown trace.
- **Behavior:** exact worker report path, `red.json`, ordered baseline/variant comparison, stdout/stderr, and any missing or incomplete evidence.
- **Closure:** `closure.json` with settings, packages, original tooling, diagnostic tooling, and native observer checks all true; zero remaining Excel; monitor/cursor proofs; immediate/delayed audit.
- **Assistance statement:** explicit confirmation whether any person, agent, OS intervention, or external process assisted teardown. The controller's own `AssistedTeardown=false` is not proof that external assistance did not occur.

Return a compact table of Quit, COM release, Collect, finalizer wait, and total wall time for baseline and variant, plus native exit relation and assertion equality. Return failures to Perplexity without implementing a fix. Never publish credentials, operational workbook contents, full warehouse paths, or unredacted screenshots/logs.

If cleanup blocks, return the exact owned process identities, last completed stage, elapsed time, evidence paths, and restoration status. Keep restoration state alive and ask for narrowly scoped assistance; do not label that run normal or release-qualified.

## Preserved contract and remaining goal

Architecture v4.11 D18-REPLAY-01 already records approval dated 2026-10-04, including explicit training-runtime execution boundaries. Plan 022 and invSys-Controls-v1.md contain matching approval. Do not reopen that approved contract based on the stale untracked guidance-025 draft, and do not invent a hybrid observation/execution contract.

No architecture amendment is introduced by this diagnostic-only tooling increment. Early Receiving B0 remains mandatory before 4be-A closes; broader replay and Admin validation-result management remain 4be-B. Existing valid evidence must be reused according to its source/package applicability.

Focused Inventory QuantityPhase GREEN remains deferred behind the shutdown investigation. The prior 140 unique checks comprise 131 PASS and nine RED; this experiment neither repeats nor closes that gate. Applicable packaged compile, layout, static maintenance, live-role, full Release 1 chain, restart/binding, and visible operator evidence remain required before their respective acceptance claims.

## Local state and exit

The diagnostic code commit is pushed as draft PR7. The original code branches and both `main` refs are unchanged. The docs review branch starts from `5557cae9b46e3a4ea7019751bd9a429542ed69fa`; only this handoff, its companion pins, and CURRENT belong to the docs change.

Preserve the pre-existing handoff-067 modification, deleted old guidance-024, and untracked guidance-023/024/025. Source-test logs and the exact publication bundle are under `invSys_fork/reports/runtime/perplexity-shutdown-isolation-source/`.

Creating a separate diagnostic worktree failed because the sandbox's Git LFS shell process could not create a signal pipe. The clean existing diagnostic checkout was switched to the review branch instead; native LFS post-checkout/post-commit commands ran without disabling hooks. No runtime or setting/security change was used to work around the sandbox boundary.

Perplexity's implementation increment stops at this test handoff, with runtime otherwise paused. Only the conditional bounded request above is authorized; 4be-A is not complete, and this is not a release-acceptance handoff.
