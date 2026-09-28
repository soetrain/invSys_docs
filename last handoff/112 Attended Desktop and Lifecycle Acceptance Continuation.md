# Attended desktop and lifecycle acceptance continuation

## Goal and release outcome

Complete invSys Release1 user acceptance under Architecture v4.11 and Plan022,
including Slice4be Event Viewer/Action Path, without regressing accepted workflows.
Active work is **4be.1 Production lifecycle observations and Action Paths** plus
its Release1 regression blockers. The goal is incomplete; its platform status was
verified **usageLimited** on2026-09-28. A new chat does not imply renewed account
allowance, architectural approval, or completion.

## Current verified state

Last verified: **2026-09-28**. Code `main` **462b0b7**, docs `main` **5ce5a5c**,
both pushed, followed by this handoff/pointer commit. Runtime correction remains
**f6d52a3**; native/presentation tooling checkpoint is **08b4add**. Code is clean.
Preserve the unrelated modified `last handoff/067 Partial Goal Action Path NAS Contract.md`
and untracked `expert guidance docs/023 Slice 4be Critique.md`; their original
hash pins still match. No Excel or diagnostic probe process remains running.

Candidate: `deploy/validation-production-paths`. Last package validation:
2026-09-25. Five builds/compiles and Operations cold start pass; compiled comparison
retains246 components with only four changed Core modules: `modConfig`,
`modProcessor`, `modEvaluationMatches`, `modEvaluationSources`. Recheck candidate
pins before the next dependent gate; do not rebuild while relevant Excel files
are open. Controls **1.255**, Plan022 and the evidence index are synchronized.

## Decisions and constraints

- **User arrangement, confirmed2026-09-28:** the user will make time to keep the
  screen unlocked for work requiring desktop access. Coordinate those gates with
  the user. Repeated attempts to remove locking have not solved the problem;
  do not continue speculative lock-policy changes or assume unattended access.
- User switched from RDP to a direct USB-C monitor and reports that the protection
  state requires Windows sign-in. Direct-monitor access currently works; an actual
  subsequent lock has not been tested. Neither connection choice nor filesystem
  access establishes access to a locked input desktop.
- User disabled **`invSys.StationUpdate`**; independently verified Disabled.
  It is the15-minute/sign-in updater that launches PowerShell visibly. Preserve
  this preference; do not silently re-enable it. `InventorySvc` is a separate
  Windows service and was not changed. The earlier administrator-prompt question
  is obsolete: the user completed the task. No agent reboot is authorized.
- Existing D5/D18 govern the runtime correction. Preserve exact immutable
  `System_Key`, unknown columns, captured-workbook binding, headless authority,
  primitive cross-package bridges and packaged Operations reuse. Do not weaken
  policy guards or restore optional Timezone to conceal provisioning on reads.
  Keep guide provenance separate from the observed run and command completion
  separate from published Designs application. New contracts require normative
  approval before implementation; a handoff cannot amend the architecture.

## Evidence and traceability

Current-candidate evidence, last run2026-09-25 unless noted:

| Gate | Verified outcome / limit |
| --- | --- |
| Lifecycle Action Paths | **708/708** functional GREEN; exact RED check set and all40 prior failures corrected. Long-run cleanup required assistance; retain that qualification. |
| Native cancellation | **94/94**, four reviewed question captures, five instrumented compiles, normal closure and preservation. Rejected first images remain recorded; capture-tool RED6/2 became GREEN8/8 before replay. |
| Paired lifecycle views | **84/84**, five reviewed images, distinct source/observed recordings, all six instructions and saved applied conclusion. Normal closure/preservation. Initial56/1 missing author-capability fixture was not product RED. |
| Lifecycle regression | **615/615**, exact preceding identities, normal closure and settings/package preservation. |
| Packaged smoke | **86/86**, normal closure, settings/packages/report restored. |
| Settings regression | **202/202**,202 unique checks, five instrumented compiles, exit0, Excel closed, settings/packages preserved at03:53:15 UTC. Reconciled from existing evidence on2026-09-28; no rerun. |
| Full chain | **5 PASS/1 harness failure**; live roles32 PASS/1 harness failure; Create Warehouse15/15. Recovery-process termination required before restoration; not GREEN. |
| Static/layout |253 components/6058 procedures/133166 lines,9 literal/45 unresolved calls,191 duplicate groups;28 caps hold,3 schemas and290 PowerShell parses pass. Prior15 source-layout checks remain applicable; nine principal native/paired images reviewed. |

**Regression blocker:** after `InventoryDomain.ProjectionRecovery.Delete`,
`modProcessor.RunBatchReportForAutomation` crashes Excel at canonical inventory
projection rebuild with0x800706BE. Windows records exceptionc0000028 in ntdll.dll.
Root cause is unresolved; those codes do not establish a driver, RDP, VBA or
processor cause. D13/Release1 require retaining the real packaged chain and
protecting projection balances/authority. A focused boundary trace was proposed
but **has not been implemented or run**. Existing
`tests/tooling/Test-Slice4beProjectionLiveControl.ps1` offers a same-session cut,
but hardcodes an older package-pin file: adapt/verify the candidate explicitly
before using it. Inject any diagnostic adapters before forms/compile, keep
packages unsaved, and emit fixed stage names only. Harness/compile failure is not RED.

**Desktop blocker,2026-09-28:** a six-minute probe records140 error5 cursor/input-
desktop failures with failed capture while the session is disconnected, followed
by39 successful samples after reconnect. A separate five-minute direct-monitor
probe records60/60 successes; no lock occurs during it. Both probes exit normally,
save no screenshots and change no lock settings. This establishes the observed
disconnect/reconnect correlation, not the cause of every historical error5.

## Do Not Repeat

- Do not rerun708 or already completed focused gates merely to recover context.
- Do not blindly repeat the unchanged full chain or label its crash product RED.
- Do not treat handoff111's host-process blockage or prepared-only native/paired
  status as current. The old exited Excel entry disappeared before03:15 UTC on
  September25; its restoration controller had already finished. No reboot occurred.
- Do not treat the latest60 access samples as acceptance under a locked desktop.
- Do not use stale prepared-test pins after the committed capture/fixture changes.
- Do not expose generated business values, credentials, settings snapshots or
  raw report Detail columns. Machine reports remain ignored by Git.

## Assumptions to re-verify

Goal usage availability, user availability for an unlocked desktop, current
desktop access, absence of Excel before gates, and candidate package pins.
The exact remaining lock trigger and the projection-crash cause remain unknown.
The five-minute monitor probe is finished; no continuing monitor is implied.

## Open questions and blockers

Current-candidate390 draft/reusable regressions and full-chain recovery remain
open. Broader55 constructed Production buttons/34 non-button semantic contracts,
other Operations/Admin coverage, guide provenance/import, carrier D5 questions,
and visible/human/NAS acceptance remain as indexed in the maintained checklist.
The earlier6.5/10 conversation estimate is not a verified acceptance metric.

## Immediate next action

When goal usage and an attended unlocked desktop are available, prepare and run a
focused same-session projection-rebuild diagnostic through the existing packaged
processor call, preserving its projection/authority assertions and package/settings
guards, before changing runtime code or retrying the full chain.

## Critical references

- Spec/plan pointers: `0 plan docs/xlam_invSys/CURRENT_SPEC.md` and
  `expert guidance docs/CURRENT.md`; controls `0 plan docs/xlam_invSys/invSys-Controls-v1.md`.
- Code evidence: `tests/integration/plan022_slice4be_remaining_acceptance.md`,
  `plan022_slice4be_production_lifecycle_paths_results.md`,
  `plan022_slice4be_production_lifecycle_visible_results.md` in the same directory.
- Boundary: `tools/validate_phase6_live_role_workflows.ps1`, step
  `Delete and rebuild canonical inventory projections`; `tools/validate_release1_full_chain.ps1`.
- Private chain evidence: `reports/runtime/production-paths-regression/chain-d41a7ecb67fb47ed95ae3c80afa029a3`.
- Private Settings evidence: `reports/runtime/production-paths-regression/settings-14c148069081439683b8237dce8b5e67`;
  count/closure receipt `reports/runtime/production-paths-settings-regression-verification.json`.
- Private desktop receipts: `reports/runtime/rdp-desktop-probe-c4eabcce85b34a88b004757dfb20aa82/summary.json`
  and `reports/runtime/rdp-desktop-probe-fd2f56d034a547ea80c9783d9820fbae/summary.json`;
  task receipt `reports/runtime/station-update-disabled-20260928.json`.
