# Handoff 265 — Guide forms verified and author onboarding decision

## Goal and release outcome

Complete 4be-A recording and guidance under D13/D18. The four reported guide-form defects are repaired in the workstation build. Broader acceptance remains open.

## Current verified state

Last verified November 24, 2026 Pacific. Code worktree `invSys_edit_receipts`, branch `work/d18-picker-save-repro`, is published through `1db8bbfa`. Docs use `work/perplexity-handoff218`, published to `work/d18-picker-save-repro-docs`; proposal commit `52184a4`. All five ordinary Excel startup entries resolve to `deploy/validation-guide-reader-reopen-20261124`, including its catalog 48. Shared deployment was untouched. The primary code checkout's edited AGENTS.md, unrelated docs edits and untracked tooling reports belong to other work. The previous PRs are closed by owner instruction.

## Decisions and constraints

The installed Admin presentation suite finished 322/322 at 15:08 Pacific,
produced 44 captures, and closed Excel normally. The final read-only check
found zero Excel processes and all five startup paths and hashes correct.
The controller stopped before starting another build or Excel test under
the workstation's required-proposal pause.

Justin authorized opening and closing Excel, managing modals and making the technical assessment. A manual intervention still requires a window capture and an attended result. First-use guide authoring is a distinct permission decision. Plan 022's November 15 checkpoint and the controls catalog explicitly require approval for ordinary onboarding. On November 24 Codex requested the decision recorded under D8-CREATE-01. The proposal is to set and confirm the initial destination Admin PIN privately during creation and support an explicit, warehouse/station-scoped Maintain Action Paths grant or revocation by a signed-in Admin. Only disposable fixture setup was previously approved. Do not put credential material in evidence.

## Evidence and traceability

| Gate | Verified result |
|---|---|
| One-click Create Guide | Before the fix, a recorded NAS source gave no feedback for 5.43 s and incurred three source reads (37/2). Feedback at 196 ms, editor at 3.11 s, two source reads on the recorded entry (40/40). Installed selected-actions route showed feedback at 176 ms and the editor at 2.58 s. |
| Reopen, Close, Cancel and X | An unloaded picker reference produced VBA error 13. A separate old reader suppressed the replacement's X (17/2). Each live form now owns its deferred close; focused gate 19/19. Installed NAS curation 99/99, including two one-click cycles per form, X and immediate reuse. No operator intervention. |
| Search and integrity | Installed picker keys took 21–26 ms and no-match 63 ms without per-key Core reads. A preceding 31-action NAS fixture took 28–83 ms. Two-version Published Guides Search took 23–26 ms, no-match 76 ms. Integrity, access loss and source preservation passed within their respective suites. |
| Regression and visible review | Installed build: five explicit compiles and cold Operations initialization, ordinary five-project/hash verification, chain 33/33 and live roles 48/48 with zero Excel crash/hang events. Preceding notice build: full reusable Production 2/2 after an earlier unexplained native crash. Preceding deferred build: Receiving B0 121/121. Codex reviewed complete and missing Receiving and Admin Save Value guide/run comparisons in How-To, Diagnostic and Compare both. Separate personal-preference walkthrough and Justin's verdict remain open. |
| Additional A1 gaps | Add Item 27/6 RED; optional notice regression 73/1; GREEN 96/96. Actual Delete Item under catalog 48 retired two entities but omitted its action and event (29/5). Catalog 49 source at `2efc5013` is unbuilt. Its expanded Delete, previous-pending, Add queue-failure and Edit Save Item tests await execution. `4352c6c0` adds two unrun picker empty-search feedback assertions; ordinary filtering itself is verified. |

## Do Not Repeat

Do not rebuild over the five installed files, treat direct fixture grants as onboarding, infer an event from an identical earlier pending payload, or count a new unrun test as behavioral RED. Preserve the earlier attended VBA error and failed Production run with their original qualifications.

## Assumptions to re-verify

Check Excel closure, the five running project paths and the current Git heads before a build. Recheck which exact candidate passed each gate. Human feedback, technical screenshot review and signed-in generated NAS tests have separate scopes.

## Open questions and blockers

Owner approval is pending for ordinary first sign-in and explicit guide grants. Remaining technical work includes actual catalog 49 qualification, Edit Save Item recording, comprehensive reachable-control accounting and final applicable guide, launcher, layout, saved-workbook restart and Receiving B0 gates on the integrated build.

## Immediate next action

After the onboarding decision, run the focused picker empty-result test against the installed build before changing its status message, then qualify the staged Delete and Edit tests on isolated packages.

## Critical references

- Architecture v4.11 D18-SEARCH-INPUT-02, D18-REPLAY-01, D8-CREATE-01 proposed onboarding and D8-USERS-01; Plan 022 and Controls v1.
- `tests/integration/plan022_slice4be_operator_review.md`, `plan022_slice4be_remaining_acceptance.md`, `plan022_slice4be_admin_inventory_entry_results.md`; `tests/tooling/Test-Slice4beConfigCommands.ps1`.
- Local-only `reports/runtime/4bea-picker-finish-20261123` contains controller exit records; NAS curation worker `b9cff875f420408faf4ae433259f992b` and Admin comparison worker `0a2d2957637346899bc60d79bde2c9e7` are under `reports/runtime/slice4be-viewer-published-read`.
