# Controller restored and Production non-button audit

## Goal and release outcome

Complete Release1/Slice4be under Architecture v4.11 and Plan022 without regressing
accepted workflows. This continuation recovers the prior test controller safely
and clarifies Production source coverage; it does not accept the unbuilt correction.

## Current verified state

Last verified: 2026-09-25,01:53 UTC. Code main **19f2cc8**, runtime correction
**f6d52a3**; docs preceding main **558e7aa**, followed by this synchronized checkpoint.
Push both and verify synchronization before ending. Preserve unrelated modified
`last handoff/067 Partial Goal Action Path NAS Contract.md` and untracked
`expert guidance docs/023 Slice 4be Critique.md`; original pins remain unchanged.

**The settings risk in handoff109 is resolved.** Its original controller now writes
closure.json at01:45:23 UTC: SettingsRestored=True, PackagesPreserved=True,
ExitCode=-1 (earlier worker termination), FullAcceptance=False. The controller,
disposable recovery test and debugger all exit. No settings values are exported.
There is no controller that must remain alive and no queued build.

**Host recovery remains pending.** Windows still enumerates the same exited Excel
process with one cleanup thread/65069 handles and increasing CPU. HasExited=True
and no live Excel process are verified; normal shutdown is not accepted. Ordinary
test guards still reject the enumerated process. No repository harness is changed
to suppress that guard. An asynchronous question requests a host restart and offers
user-performed restart, explicit authorization for agent restart, or keeping the PC
running/pausing Excel tests. No response yet; do not infer restart authorization.
The earlier unlocked/RDP question is superseded by this recovery state.

Final `deploy/validation-production-paths` is **not built/compiled/GREEN**. Intermediate
`deploy/validation-production-policy-read` proves D5 preservation and evaluator RED.
Frozen baseline remains `deploy/validation-production-lifecycle`. No new Excel has
been started. Controls1.252, coverage audit1.4 and Plan022 agree; Architecture and
runtime/test source are unchanged during this continuation.

## Decisions and constraints

D5/D18 already govern the pending source correction: processor reads existing
Config without provisioning, while lifecycle evaluation distinguishes exact command
completion from published Designs application. See handoff109 and the primary
evidence report for implementation details. Preserve five RED test hashes for GREEN.

The new non-button review is source accounting, not approved new IDs/outcomes or
runtime behavior. Actual Output/Batch Note Change handlers update reusable memory;
quantity-mode choices can clear editor fields; hidden connection/list and mirrored
List/Tree callbacks can duplicate a deliberate action; non-reusable location changes
can clear captured staging allocation cells. Do not flatten these into read-only
Navigation. D18 prohibits keystroke/entered-value capture. Explicit semantic commit
and owner facts still need packaged tests before implementation. No legacy fallback
or identity exception is authorized. Catalog13 retains13 registered Production IDs;
55 constructed buttons remain pending.

## Evidence and traceability

Code `tests/integration/plan022_slice4be_production_lifecycle_paths_results.md` retains
complete **668 PASS/40 expected FAIL across708 unique checks**, real24 action pairs,
recording/publication/detail identity and read-only preservation. GREEN is pending.
Private report: `reports/runtime/slice4be-production-lifecycle-paths/fdbe45d1bed1455d8ae27457c3cc2e5c/red.json`.
Controller: `reports/runtime/production-lifecycle-paths-controller/02aec33d88644412bff06fcf6b1dfa2a`.
Its closure receipt must be read alongside assisted-cleanup and recovery evidence.

Recovery first uses a disposable waiting controller: temporary process-local
Get-Process filtering retains live processes and excludes HasExited=True entries;
in-memory canary survives. Supported PowerShell Enter-PSHostProcess/Debug-Runspace
then applies that wrapper only within the original controller's paused wait loop.
Its unchanged restoration/comparison finishes and all processes involved exit.
No snapshot values/credentials are dumped or written. ExcelClosed=True in this
recovered receipt means no live Excel, not disappearance of the retained OS entry.
Private receipts: `test-exited-controller-recovery-result.json` and
`production-paths-controller-recovery.json` under reports/runtime.

The source audit matches all34 non-button handler identities exactly, reviews owner
helpers and corrects stale SubmitProcessAction/SubmitRecipeAction/seven-ID wording.
Receipt: `reports/runtime/production-nonbutton-source-review.json`. No new behavior
means no manufactured D13 RED; validation is source reconciliation, diff/link checks
and preservation pins. Prior final source static remains253/6058/133166 lines,
9 literal/45 unresolved calls,191 duplicate groups,28 caps,3 schemas,285 PS parses,
15 layout checks. It does not substitute for package compile or regressions.

## Do Not Repeat

- Do not preserve/wait on the old controller: it has completed restoration and exited.
- Do not treat assisted recovery or HasExited as normal shutdown acceptance.
- Do not start another Excel gate or authorize a host restart by elapsed time alone.
- Do not alter the five protecting tests to make evaluator GREEN easier.
- In PS5, assign ConvertFrom-Json arrays before foreach; piping the resulting array
  directly can compare arrays of hashes incorrectly. Explicit per-file recheck
  proves all five pins unchanged; the initial pipeline comparison was a false alarm.

## Assumptions to re-verify

Host restart response, OS process state and desktop input are pending/stale. Cursor
probe briefly succeeded at01:10 UTC and returned error5 at01:29. Group Policy and a
shared cause with the shutdown/RPC failures remain unproven. No policy changed.
Recheck package/test hashes, unrelated-doc pins and Git state after any restart.

## Open questions and blockers

Host recovery blocks normal Excel gates. Final package compile/708 GREEN,615 lifecycle,
390 draft,202 Settings,86 smoke, full-chain/live-role and reusable regressions remain.
Native cancellation, distinct-source/observed-run How-To/Diagnostic comparison,
broader Operations/Admin controls and visible/human/NAS acceptance remain open.
Earlier chain/reusable RPC800706BE failures retain their candidate-specific evidence.

## Immediate next action

Resolve the pending restart request, verify Excel is absent and desktop ready, then
build the five final XLAMs and run compile followed by the unchanged708-check GREEN.

## Critical references

`tests/integration/plan022_slice4be_remaining_acceptance.md`; Production coverage
audit1.4; `reports/runtime/production-paths-test-pins.json`; handoff109's source/build
commands. Use `deploy/validation-production-paths` explicitly for compile and
`Test-Slice4beProductionLifecycle.ps1 -ActionPaths -Phase GREEN`. Private sequential
regression runner: `reports/runtime/run-production-paths-regression.ps1`.
