# Lifecycle Action Path RED and pending Windows cleanup

## Goal and release outcome

Complete Release1/Slice4be under Architecture v4.11 and Plan022. This checkpoint
advances existing D5 read-only Config and D18 lifecycle recording/diagnostic rules.
It is unaccepted work in progress; do not call the goal or this correction complete.

## Current verified state

Last verified: 2026-09-25,01:40 UTC. Code main **f6d52a3**; docs main preceding
checkpoint **d6aabc5**, followed by this synchronized record/pointer commit.
Preserve unrelated modified `last handoff/067 Partial Goal Action Path NAS Contract.md`
and untracked `expert guidance docs/023 Slice 4be Critique.md`; original pins remain
in code `reports/runtime/evaluation-selection-unrelated-doc-pins.json`.

**Windows cleanup is unresolved.** Controller
`reports/runtime/production-lifecycle-paths-controller/02aec33d88644412bff06fcf6b1dfa2a`
retains the original settings snapshot in memory. Do not terminate that controller.
Its worker saves all708 results and reaches terminal output, but shutdown stalls.
Exact owned Excel and then completed worker termination are requested. Windows
reports Excel HasExited=True/no window, yet still enumerates one thread/65069
handles with increasing CPU; taskkill reports no running instance. No closure.json
exists yet. Final outer settings restoration/package preservation are **pending**.
The unstarted build waiter was stopped; no build will start automatically.
Tool session65047 owns the surviving controller. No other Excel gate may start
before cleanup and restoration are verified. Preserve its in-memory state across
continuation; avoid restarting the host while restoration is unresolved.

Frozen baseline: `deploy/validation-production-lifecycle`. Intermediate five-package
candidate: `deploy/validation-production-policy-read`, D5 fix plus old evaluator.
Final `deploy/validation-production-paths` is **not built**. Source evaluator/helper
changes have no final packaged compile or GREEN. Controls1.251/coverage1.3/Plan022
record this explicitly; Architecture is unchanged because the rules already exist.

## Decisions and constraints

The new actual Process Save recording exposes processor read-side provisioning:
absent optional Timezone makes OpenOrCreateConfigWorkbookRuntime normalize sheets,
delete tracking policy and save Config. Completion correctly rejects changed policy.
Do not weaken that guard or repair the optional column to mask this D5 breach.
Use Core's existing ResolveExistingConfigForRead through a direct typed Core-local
call, preserve borrowed workbooks, close only owned transient reads without saving.
The obsolete private name-only lookup is removed after reviewed reachability,
five intermediate compiles and24 preserving action cases.

D18 evaluator mappings accept CONFIRMED for exactly six lifecycle controls.
Designs application needs exact nonempty Business-event lines/outcomes, EventID,
WarehouseId, AppliedAtUTC and positive integer AppliedSeq (published as text),
available coverage, full line count/hash and empty SystemKeys. Missing Submitted
IDs are Awaiting only under complete available coverage; uncertain/malformed/
unavailable evidence is Incomplete. Inventory semantics stay protected. Read-only
evaluation must preserve original records, journal, saved authority and publication.

## Evidence and traceability

Primary code report: `tests/integration/plan022_slice4be_production_lifecycle_paths_results.md`.
D5 RED controller e442f999bc55459e92b1289a33764051, report
`slice4be-production-lifecycle-paths/324aa20923d74262b6e95db23104a66b/red.json`:
144 PASS/11 FAIL, including two actual D5 failures, eight evaluator assertions and
one prerequisite-stop marker (not behavioral RED). Normal cleanup/settings restore.

Complete intermediate evaluator RED: **668 PASS/40 expected FAIL**,708 unique IDs,
report `slice4be-production-lifecycle-paths/fdbe45d1bed1455d8ae27457c3cc2e5c/red.json`.
Eight supplemental evidence failures, six confirmed-command conclusions, twelve
confirmed-source checks, twelve pending-source checks and two ordered series fail.
All24 actual action pairs,50-entry journal, real Admin publication, Viewer detail,
matched original occurrences, read-only preservation and D5 policy/bytes pass.
Five instrumented compiles pass. No harness-result failures; assisted unresolved
shutdown remains separate. All report/controller paths above are under reports/runtime.
Five test files are pinned in `production-paths-test-pins.json`; retain exact tests
for final GREEN. `production-paths-red-verification.json` records identities/counts.

Final source static: `production-paths-final-static`;253 components/6058 procedures/
133166 lines (+12),9 literal/45 unresolved calls unchanged,191 duplicate groups
(one fewer), all28 caps held. Three schemas,285 PS parses and15 source-layout
checks pass (`production-paths-static-verification.json`). This cannot replace GREEN.
Prior baseline615/615 lifecycle,390/390 draft paths,202 Settings and86 smoke are
not current-source regressions. Prior chain/reusable RPC failures remain open.

## Do Not Repeat

- Do not rebuild with an unresolved Excel/controller lifecycle or lose the original
  settings snapshot. No second COM attachment after Quit; no unrelated process kills.
- Do not treat the worker's final output or HasExited alone as completed restoration.
- Do not call the unbuilt source correction GREEN or accept the intermediate package.
- Two FAILED original occurrences require two ordered expected steps to select the
  AfterAppend record instead of the earlier BeforeAppend record.
- Hash open fixture workbook bytes with FileShare.ReadWrite.
- Keep the older `plan022_slice4be_production_paths_results.md` unchanged; the new
  report includes **lifecycle** in its name. A temporary overwrite was restored.

## Assumptions to re-verify

Cursor access briefly succeeds at01:10 UTC, then returns error5 again at01:29.
Group Policy remains unproven; no Windows settings changed. An asynchronous
question asks the user to keep host/client visibly unlocked and RDP open. No new
native captures are accepted. Recheck desktop, process/controller state, packages,
pins and Git synchronization. Distinguish late successful cleanup from this checkpoint.

## Open questions and blockers

Windows cleanup currently blocks building/testing. After recovery: final compile,
708 GREEN,615 lifecycle/390 draft/202 Settings/86 smoke regressions, full chain and
reusable Production. Native lifecycle cancellation, How-To/Diagnostic comparison,
broader55 pending constructed Production buttons, other role coverage and visible/
human/NAS acceptance remain open. Do not infer a common cause for error5, shutdown
or earlier RPC800706BE failures. No architecture change has been requested here.

## Immediate next action

Inspect the retained controller's closure.json and owned-process recovery receipts;
only after verified closure/restoration build the five final packages, compile and
rerun the unchanged `Test-Slice4beProductionLifecycle.ps1 -ActionPaths -Phase GREEN`.

## Critical references

Build: `tools/build-xlam.ps1 -RepoRoot . -OutputRoot deploy/validation-production-paths
-Projects @('Core','InventoryDomain','DesignsDomain','Operations','Admin') -Apply`.
Compile: `tests/tooling/Test-PackagedVbaCompile.ps1`, SourceReportPath
`reports/runtime/production-paths-compiled.json`. GREEN explicitly selects final
DeployRoot. Private `reports/runtime/run-production-paths-regression.ps1` prepares
sequential Lifecycle/Draft/Settings/Smoke/Chain/Reusable gates and preserves settings,
packages and tracked reports; inspect shutdown evidence rather than assuming it.
Native UOM dialog observer pattern: `tests/tooling/Slice4beAdminUomProbe.ps1`.
Current lifecycle adapters bypass confirmations; native cancellation requires its
own actual-handler test. Guide source and observed run must remain explicit and
distinct; do not substitute pure source-evidence tests for either presentation.
