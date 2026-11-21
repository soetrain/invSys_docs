# Handoff 229: v33 Settings host assertion and fact probe

Date: 2026-11-12 Pacific. Predecessor: handoff 228.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. The current Settings gate remains unqualified after a test-host assertion failed before its 202 behavioral checks.

## Current verified state

Last verified 2026-11-12: code `main` is `6d247543` with Justin's unrelated `AGENTS.md` edit untouched. Draft code PR14 `d94768c7` is clean in its own `work/d18-settings-saved-host` worktree, stacked on PR13. The execution worker was restored clean to `work/d8-inv-edit-receipts` at `550d9ff0`. Docs `work/perplexity-handoff218` has Plan 022 and Controls **1.741** at `0685835`; this handoff follows. Unrelated docs deletion, edit and untracked files remain untouched. Excel is closed. Recheck before any runtime.

## Decisions and constraints

- Architecture v4.11's saved-workbook operator path is unchanged. PR14 modifies only test tooling; the five candidate XLAMs, source VBA, ControlIds and wording remain unchanged.
- The separate owner approval for one PR4-BASE-09 v33 run was used. No v33 retry or reuse of that authorization. A new, distinct fact probe is prepared in a copied worker but has not been authorized or dispatched.
- Preserve natural Excel closure, pre-run settings, 539 pinned inputs, unrelated files, and the 2026-11-12 **21:08 UTC** shift boundary. Stop at lockscreen, verified input failure or required proposal approval.

## Evidence and traceability

- `settings-policy-v33-pr4-base-09-20261112`: checks-only passed 539 pins/202 expected identities. The sole approved run passed six instrumented compile/probe checks, then failed `Harness.Exception.Admin-generated fixtures` because the test's combined saved-host retention assertion failed after fixture creation. Worker exit 1; 196 expected checks missing; no Settings behavioral RED or GREEN. Both native exits verified, stdout quiet→Excel exit 1.92 s, settings restored, inputs preserved, no process termination, delayed Application audit zero Excel faults. No modal. Dispatch and closure are in that local folder.
- The failed assertion combines exact host path, `Workbook.Saved` and workbook count, without recording which changed. The earlier diagnostic saved-host probe was created **after** fixtures and stayed near 927–935 GDI; it does not settle the **before-fixture** v33 state. Do not remove a guard or attribute fault to product code from these facts alone.
- `4be-settings-host-state-probe-20261112` has a 544-file copied test toolset; only the copied Settings worker differs from clean PR14, SHA256 `82D88A2D...`, parser zero errors. It records six nonsensitive host facts after fixtures and returns before Settings behavior. No run or approval exists. Its `codex-host-state-probe.md` specifies the one-run scope.
- Earlier independent Inventory, Recording and technical B0 replay results remain in handoff 228/Plan 022. `B0Accepted=false`; broad 4be-A and Release 1 acceptance remain open.

## Do Not Repeat

Do not rerun v32/v32.1/v33, treat six compile checks as Settings GREEN, change the product to accommodate an unclassified harness assertion, or broaden the copied probe into a second Settings gate. The fixed v33 start window is consumed.

## Assumptions to Re-verify

Recheck PR heads, branch cleanliness, package hashes, Excel closure, desktop/input and available resources. The copied fact probe has not been run and may itself expose a fixture prerequisite rather than the original retention fact.

## Open questions and blockers

Which saved-host condition changes during fixture creation; fresh Settings packaged GREEN; remaining Event Viewer/detail, guide transfer, other Admin/Operations families, retained-formatting, human B0, live-role, restart/binding, full-chain and visible operator acceptance. Residual Execute slowness and the Microsoft-reported Excel exit defect remain separate.

## Immediate next action

Obtain separate one-run approval for the copied host-state probe, run it once with unchanged five XLAMs, and use its six facts to make the smallest test-harness correction before repinning any acceptance run.

## Critical references

Architecture v4.11 operator deployment model and D18; Plan 022 PR4-BASE-09 checkpoint; Controls 1.741; draft code PR14; docs PR5; `tests/tooling/Test-Slice4beConfigCommands.ps1`; local v33 and host-state probe folders above.
