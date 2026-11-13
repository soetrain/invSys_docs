# Handoff 231: v34 Settings worker pass and restart controller gate

Date: 2026-11-12 Pacific. Predecessor: handoff 230.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. The packaged Settings worker passed its 202 checks; controller qualification and broad acceptance remain open.

## Current verified state

Last verified 2026-11-12: code `main` is `6d247543` with Justin's unrelated `AGENTS.md` change untouched. Draft code PR14 is clean at `b8f94241` on `work/d18-settings-saved-host`. Execution worker `C:\Users\Justin\repos\invSys_edit_receipts` is clean and detached at `b8f94241` for v35; return it to `work/d8-inv-edit-receipts` (`550d9ff0`) after the run or if the run is abandoned. Docs PR5 is on `work/perplexity-handoff218`; Plan 022 / Controls **1.743** record v34. Other user docs changes remain untouched. Excel is closed. The owner's shift ends **21:08 UTC**.

## Decisions and constraints

- Architecture v4.11's saved operator workbook contract is unchanged. PR14 modifies only test tooling; the five validation XLAMs and product contracts are unchanged.
- v34's one-run approval is consumed; do not retry. v35 received approval after its **19:55 UTC** cutoff and never dispatched. Its pinned controller was not changed. v36 is a separate controller/window and needs separate approval before **20:15 UTC**. No assisted Excel close, process kill, or retry.
- The `soetrain` docs branch `work/perplexity-handoff218` points to an older merged PR6 tree. Docs PR5's actual head is the `justinwj` fork branch; push that fork branch, and do not merge the older `soetrain` branch into current docs.

## Evidence and traceability

- v34 worker: 202/202 exact named PASS, zero FAIL, exit 0; `Preference.NewExcelProcessRestoresSavedChoice` passed. The pinned fixture naturally closed first Excel and started one replacement. v34 controller rejected replacement PID 16096 as extra, so its gate failed. Settings and all 539 pins were preserved; both Excel processes ended, no termination; delayed Application audit found zero Excel faults. `run-S1/closure.json` and `series-closure.json` are in `reports/runtime/settings-policy-v34-pr4-base-10-20261112`.
- v35 changed only the ignored runtime controller: exactly one replacement after first natural exit, pinned Preference PASS and unassisted restart proof, native exit of both Excel lifetimes, no further processes. Parser and checks-only passed without Excel, but the approved start window expired before dispatch. v36 preserves that rule and changes only series identity/window. Manifest `68338D11...`, expected identities `E73F7953...`, launcher `22FCBD39...`; parser zero errors; checks-only passed 539 pins / 202 identities without Excel. Exact hashes, scope and command are in `reports/runtime/settings-policy-v36-pr4-base-10-20261112/codex-settings-policy-v36.md`.
- Plan 022 and Controls 1.743 are committed and pushed on docs PR5's fork branch. v34 worker GREEN is not qualified Settings GREEN, 4be-A or Release 1 acceptance.

## Do Not Repeat

Do not rerun v32-v34, dispatch expired v35, treat v34's controller failure as product RED, or treat the second Excel as an unrelated process without checking the Preference fixture. Do not merge the older `soetrain` docs working branch into the newer PR5 head.

## Assumptions to Re-verify

Recheck owner approval, v36 time window, clean detached worker, no Excel, desktop/input, five XLAM hashes, 539 pins, launcher hash and no v36 dispatch before any run.

## Open questions and blockers

v36 approval/qualified Settings gate; remaining Event Viewer/detail, guide transfer, other Admin/Operations, retained formatting, human Receiving B0, live role, full chain and visible operator acceptance. Execute slowdown and the Microsoft-reported Excel exit defect remain separate.

## Immediate next action

If fresh approval arrives before 20:15 UTC, rerun v36 checks-only and dispatch its single approved run; otherwise leave runtime paused and prepare a new bounded window.

## Critical references

Architecture v4.11 D18; Plan 022 PR4-BASE-10 checkpoint; Controls 1.743; draft code PR14; docs PR5; `tests/tooling/Test-Slice4beConfigCommands.ps1`; v34/v35/v36 folders above.
