# Handoff 232: v36 Settings GREEN and Event Viewer detail gate correction

Date: 2026-11-12 Pacific. Predecessor: handoff 231.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. Settings now has current-package GREEN; Event Viewer detail and broader acceptance remain open.

## Current verified state

Last verified 2026-11-12: code `main` is `6d247543` with Justin's unrelated `AGENTS.md` change untouched. Draft code PR14 is clean at `b8f94241`. Execution worker `C:\Users\Justin\repos\invSys_edit_receipts` is clean and detached at `b8f94241` for the prepared v39 check; restore `work/d8-inv-edit-receipts` (`550d9ff0`) after the run or if abandoned. Docs PR5 branch `work/perplexity-handoff218` has Plan 022 and Controls **1.744**. Other user docs changes remain untouched. Excel is closed. The owner's shift ends **21:08 UTC**.

## Decisions and constraints

- The five validation XLAMs and product contracts are unchanged. v36's one-run approval is consumed. v35 never dispatched because its fixed cutoff passed before approval.
- v37 approval was consumed by one dispatch; its controller passed an unsupported compile switch and stopped before Excel. v38 approval was consumed by one dispatch; its worker passed 90/90, but the controller pinned only 85 names. Do not retry either. A distinct v39 retains the supported worker command and pins all 90 observed identities. It needs new approval before **20:50 UTC / 12:50 PM Pacific**. No retry, assisted Excel close, process kill or input rebinding.
- The `soetrain` docs branch `work/perplexity-handoff218` points to an older PR6 tree. Docs PR5's current head is the `justinwj` fork branch. Do not merge the older branch into the current docs work.

## Evidence and traceability

- v36 PR4-BASE-10 visible packaged Settings/policy/detail/preference: **202/202 exact PASS**, worker exit 0, controller `StopReason=null`, restart proof and both Excel native exits verified, no extra Excel, settings restored, 539 pins preserved, no process termination, no modal intervention, delayed audit zero Excel faults. Run 6.49 minutes on Excel 16.0.20527.20196. The negative stdout-quiet-to-first-Excel-exit metric is inapplicable because the first Excel exits before later checks. Evidence: `reports/runtime/settings-policy-v36-pr4-base-10-20261112`.
- v37: worker exited 1 with the exact unsupported-compile-switch refusal; zero stdout/behavioral checks, no Excel observed, settings/pins preserved, no process termination, delayed audit zero Excel faults. It is harness setup failure, not product RED. Evidence: `reports/runtime/viewer-detail-v37-pr4-base-10-20261112`.
- v38: the fixture itself compiled five packages and printed the historical five `Harness.InstrumentedCompile.*` checks. Worker passed 90/90 with exit 0; Excel exited naturally, settings/inputs preserved, no intervention, delayed audit zero faults. Controller expected 85 and stopped on five unexpected names, so this is worker GREEN but not a qualified controller gate. Read-only comparison proved exactly 90 unique PASS names matching v39's expected file, with zero missing or unexpected names and one GREEN summary. Evidence: `reports/runtime/viewer-detail-v38-pr4-base-10-20261112`.
- v39 is test-only in ignored runtime, with the same supported worker switches and 90 exact identities. It exercises selected detail, multiline/long text, native cursor scroll/typing and screenshots on unchanged five XLAMs. Manifest `68338D11...`, expected `F1B24DC6...`, launcher `65FC6248...`; parser zero errors, checks-only passed 539 pins / 90 identities without Excel. Exact hashes, scope and command: `reports/runtime/viewer-detail-v39-pr4-base-10-20261112/codex-viewer-detail-v39.md`. No v39 dispatch or authorization file exists.
- Plan 022 and Controls 1.744 are committed and pushed on docs PR5's fork head. v36 is standalone Settings evidence, not 4be-A or Release 1 acceptance.

## Do Not Repeat

Do not rerun v32-v38, dispatch expired v35, treat v34/v37/v38 controller failures as product RED, or use the erroneous 1:45 PM translation. Do not merge the older `soetrain` docs working branch.

## Assumptions to Re-verify

Recheck v39 approval, cutoff, clean detached worker, no Excel, desktop/cursor readiness, five XLAM hashes, 539 pins, launcher hash and no v39 dispatch before the run.

## Open questions and blockers

v39 current-package Event Detail controller evidence; Viewer published reads/groups/refresh, guide transfer, other Admin/Operations, retained formatting, broad Inventory, human Receiving B0, live role, full chain and visible operator acceptance. Execute slowdown and Microsoft-reported Excel exit defect remain separate.

## Immediate next action

If approval arrives before **20:50 UTC**, rerun v39 checks-only and dispatch one hands-off visible Event Detail run; otherwise leave runtime paused and restore the worker branch.

## Critical references

Architecture v4.11 D18 and Event Detail scrolling decision; Plan 022 PR4-BASE-10 checkpoint; Controls 1.744; code PR14; docs PR5; `tests/tooling/Slice4beViewerEventDetail.ps1`; v36/v37/v38/v39 folders above.
