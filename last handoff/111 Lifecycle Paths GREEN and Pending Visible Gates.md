# Lifecycle paths GREEN and pending visible gates

## Goal and release outcome

Complete Release1/Slice4be under Architecture v4.11 and Plan022 without regressing
accepted workflows. This checkpoint proves D5-preserving Production lifecycle
recording and D18 command/applied/pending diagnostics; full acceptance remains open.

## Current verified state

Last verified: 2026-09-25,03:01 UTC. Code main **52d6a47**, runtime correction
**f6d52a3**; docs main **6ab7d3e**, followed by this handoff commit. Both checkpoint
commits are pushed. Code is clean. Preserve unrelated modified
`last handoff/067 Partial Goal Action Path NAS Contract.md` and untracked
`expert guidance docs/023 Slice 4be Critique.md`; their original pins still match.

Active slice: **4be.1 lifecycle observations and their Action Paths**. Final
`deploy/validation-production-paths` is built; Operations cold start and all five
package compiles pass. Compiled comparison retains246 components and changes only
Core `modConfig`, `modProcessor`, `modEvaluationMatches`, `modEvaluationSources`.
The unchanged protecting gate passes **708/708**, zero harness failures, the exact
RED check set, all40 prior failures corrected and five test hashes preserved through
the GREEN run. Two runner files were extended only after that verification.

**Settings/controller recovery is complete.** GREEN controller closure at02:52:53
verifies SettingsRestored=True and PackagesPreserved=True. Its worker needed
termination after saving terminal results; ExitCode=-1 records that assistance.
No Excel termination was requested in this GREEN run. Controller and debugger exit;
there is no in-memory snapshot controller that must remain alive.

**Host cleanup still blocks Excel gates.** Windows reports the same Excel entry
HasExited=True while retaining one kernel-busy thread/65024 handles. Ordinary guards
still reject that entry. Do not start another Excel gate until it disappears.
The latest cursor probe succeeds at03:00 UTC/error0. No reboot or Windows setting
change is performed by the agent. An asynchronous question offers waiting,
user-performed restart, or explicit authorization for an agent restart; no answer
has arrived. Never infer reboot authorization from the user's permission to close Excel.

Controls1.253 and Plan022 agree. Architecture and runtime source are unchanged in
this continuation. Refreshed static metrics remain253/6058/133166 lines,
9 literal/45 unresolved calls,191 duplicate groups;28 caps,3 schemas and289 PS
parses pass. Earlier15 layout checks remain applicable to unchanged form source.

## Decisions and constraints

Existing D5/D18 govern the four-module correction; no new contract is introduced.
Keep exact System_Key identity, unknown columns, captured workbook, headless
authority, primitive bridges and packaged Operations reuse. Do not weaken the
completion policy guard or add optional Timezone to conceal provisioning on reads.
Distinguish command completion from published Designs application; Designs carries
exact owning references and applied lines without inventing inventory SystemKeys.

The new native and presentation helpers are **prepared, not packaged-tested**.
Native cancellation invokes four real Release/Obsolete handlers with confirmation
enabled, scopes the observer to the generated process and exact questions, captures
them and clicks only No. It checks CANCELLED records, default No, source/draft
preservation and actual captured-workbook identity despite an active decoy.
The presentation gate records two independent six-action series, authors a guide
from one, explicitly evaluates the other and compares How-To/Diagnostic/Compare both.
Five expected captures include the final instruction and saved conclusion.
All adapters are installed into unsaved packages before forms/compile.

## Evidence and traceability

Primary report: code `tests/integration/plan022_slice4be_production_lifecycle_paths_results.md`.
GREEN result:
`reports/runtime/slice4be-production-lifecycle-paths/df9129b233fd490b977db6aff397facf/green.json`.
Controller:
`reports/runtime/production-lifecycle-paths-controller/21b2eeea63354e18a4d8973e03c96b31`.
Receipts: `production-paths-green-verification.json`, `production-paths-green-recovery.json`,
`production-paths-compiled-comparison.json`, `production-paths-visible-static-verification.json`.
New prepared tests are pinned in `production-lifecycle-visible-prepared-pins.json`.

Cleanup delays the final general checks until about02:47:53, then all708 results
save. Read-only samples show growing Section/Event handles during the long gate,
broadly steady GDI/USER counts, and predominantly kernel CPU during cleanup.
Application/Display queries find no corresponding failure. Root cause is unproven;
do not assert GPU, RDP, Group Policy, throttling or runtime code as the cause.
The already proven process-local Get-Process filter excludes HasExited entries and
preserves live processes only inside the original restoration controller. Its
unchanged restoration/comparison completes; no settings values are exported and
no repository cleanup guard is relaxed. ExcelClosed in closure means no live Excel,
not disappearance of the retained Windows entry or normal shutdown acceptance.

## Do Not Repeat

- Do not rerun708 merely to rediscover functional GREEN; its exact preserved proof exists.
- Do not wait on or recreate the old settings controller; it has restored and exited.
- Do not label prepared tests, offline C# compile, or assisted cleanup as visible acceptance.
- Do not blindly retry earlier full-chain/reusable RPC failures; retain their candidate scope.
- Assign ConvertFrom-Json arrays before foreach in PS5; avoid false array hash comparisons.

## Assumptions to re-verify

Check Windows process disappearance, pending restart answer and desktop access.
Recheck package and unrelated-document pins after host recovery. New native and
presentation fixtures may expose harness issues; compile/fixture failures are not
behavioral RED. Inspect all generated images before accepting visible evidence.

## Open questions and blockers

Host cleanup blocks native cancellation, paired lifecycle presentations and current
candidate615 lifecycle/390 draft/202 Settings/86 smoke/full-chain/live-role/reusable
regressions. Normal shutdown remains unaccepted. Broader55 constructed Production
buttons,34 non-button semantic contracts, other Operations/Admin coverage, guide
provenance/import, carrier D5 questions and visible/human/NAS acceptance remain open.
See `tests/integration/plan022_slice4be_remaining_acceptance.md`; the goal is not complete.

## Immediate next action

Once Excel is absent and desktop access succeeds, run
`Test-Slice4beProductionLifecycle.ps1 -DeployRoot deploy/validation-production-paths -Phase GREEN -NativeCancellation`,
inspect all four images, then run the separate `-Presentation` route.

## Critical references

`tests/tooling/Slice4beProductionLifecycleNative.ps1`, `Slice4beProductionNativeChoice.ps1`,
`Slice4beProductionLifecyclePresentation.ps1`; shared `Test-Slice4beConfigCommands.ps1`;
private sequential regression runner `reports/runtime/run-production-paths-regression.ps1`.
Use the final deployment explicitly. Do not rebuild unchanged packages to install
test probes; the gate injects unsaved adapters into its disposable package copies.
