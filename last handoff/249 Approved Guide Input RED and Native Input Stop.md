# Handoff 249: approved guide input RED and native input stop

Date: 2026-11-17 Pacific. Predecessor: handoff 248.

## Goal and release outcome

Complete Slice 4be-A under Architecture v4.11 D18-REPLAY-01 and Plan 022 without regressing Release 1. Approved D18-GUIDE-INPUT-01 addresses the operator's slow guide typing; its implementation and acceptance remain open.

## Current verified state

Last verified 2026-11-17: code worker `C:\Users\Justin\repos\invSys_edit_receipts` is `work/d18-picker-save-repro` at `6fabc2af`, pushed to `soetrain/invSys_fork` draft PR17, base `work/d14-packaged-chain-coverage`. Only pre-existing generated `tests/tooling/reports/` remains untracked. Docs `C:\Users\Justin\repos\invSys_docs` is `work/perplexity-handoff218`; this handoff and the Plan/Controls checkpoint follow normative approval commit `0b47aad` on draft PR8, base `work/excel-shutdown-timing`. Preserve unrelated dirty and untracked docs. Main code checkout remains `main` at `6d247543`, with Justin's unrelated `AGENTS.md` edit.

The last test-owned Excel PID 43860 exited naturally after the worker. Confirm Excel count zero and candidate hashes before another test or build. No product form, Core, Domain or XLAM change was made in this session.

## Decisions and constraints

Justin approved D18-GUIDE-INPUT-01 on 2026-11-17. Architecture v4.11, Plan 022 and Controls 1.767 agree: local text buffers for four guide fields; guarded Core flush before dependent actions and ordinary blur; previous editing StepId owns dirty step text; actual Cancel/X must discard without a `WriteText` attempt or publication even with stale authority. Do not treat the test-only Core write counter as a product change. Stop desktop testing after verified mouse/input failure under `AGENTS.md`.

## Evidence and traceability

- Focused packaged actual-handler RED on unchanged five XLAMs: **28 PASS / 12 behavioral FAIL** plus one native harness exception. Instrumented compile and disposable fixtures passed. Per-character Core writes and slow name/tags/instructions/step text remained; ordinary blur and dirty step-switch lacked the required single flush. Exact text and prior StepId content, two immutable Save routes, and source bytes passed. Ignored raw report: worker `reports/runtime/slice4be-viewer-published-read/188b501884294859b661733fe7a5ce32/` (`red.json`, `guide-editor-input-timing.json`). Timings for 20/16/25/22 characters: 42.18/34.01/54.03/48.00 seconds, all above 2.5-second bounds.
- The native Cancel/X proof is **unqualified**. First attempt could not find the dynamically created Cancel button through UI Automation; second could not foreground the form; third reached native input but failed cursor restoration (`SetCursorPos=False`, immediately reported Win32 error 0). The third run reported 28 PASS / 13 FAIL including the harness exception. No X result was reached. This is verified input failure, not a behavioral RED or GREEN. The failed worker's Excel later exited naturally.
- The likely slow path remains `frmActionPathGuide.StageText` -> Core `modActionGuideDraft.WriteText` -> guarded source/policy/Auth reads. This is source-supported, not a timed attribution of each guard. Governing rule: Architecture D18-GUIDE-INPUT-01 and D13 packaged handler evidence.

## Do Not Repeat

Do not count the native harness exception as product RED, the text-preservation passes as responsiveness GREEN, or local picker timing as NAS usability acceptance. Do not rerun desktop input or Excel until the cursor works again. Do not delete the policy-blocked generated `tests/tooling/reports/` or disturb unrelated docs edits.

## Assumptions to Re-verify

Recheck the active input desktop and cursor movement/restore, Excel closure, five candidate hashes, worktree status and PR heads. The native form geometry and foreground path in `Slice4beGuideInputNative.ps1` have not qualified.

## Open questions and blockers

Native Cancel/X zero-write ordering, dirty reorder/removal, expectation entry/return, refresh/rebind, stale policy/capability/session refusal, packaged GREEN, compile/layout/static/live-role/full-chain/restart and visible NAS usability remain open. 4be-A, 4be-B and Release 1 acceptance remain open. The Windows input failure stopped this shift early.

## Immediate next action

After independently verifying native cursor movement and restoration, repair the disposable Cancel/X test path and obtain qualified packaged RED for the remaining event-ordering cases before editing `frmActionPathGuide`.

## Critical references

Architecture v4.11 `D18-GUIDE-INPUT-01`; Plan 022; Controls 1.767; code PR17; docs PR8; `tests/tooling/Slice4beGuidePickerStoppedRun.ps1`; `tests/tooling/Slice4beGuideInputNative.ps1`; `tests/tooling/Slice4beGuideDraft.ps1`; `frmActionPathGuide`; `modActionGuideDraft.WriteText`; final candidate `reports/runtime/empty-holds-green2-20261116/addins/`; source pin `reports/runtime/4be-a-human-review-20261116b/candidate.json`.
