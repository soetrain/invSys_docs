# Handoff 230: PR4-BASE-10 Settings gate prepared

Date: 2026-11-12 Pacific. Predecessor: handoff 229.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. The saved-workbook Settings test host is corrected, but its full 202-check packaged gate and broad acceptance remain open.

## Current verified state

Last verified 2026-11-12: code `main` is `6d247543` with Justin's unrelated `AGENTS.md` edit untouched. Draft code PR14 is `b8f94241`, clean in `work/d18-settings-saved-host`, stacked on PR13. Execution worker `C:\Users\Justin\repos\invSys_edit_receipts` is clean and detached at `b8f94241` for the prepared gate; return it to `work/d8-inv-edit-receipts` after the run or at session close. Docs `work/perplexity-handoff218` has Plan 022 / Controls **1.742** at `020a06d`; this handoff follows. Unrelated docs changes remain untouched. Excel is closed.

## Decisions and constraints

- Architecture v4.11's saved `.xlsm`/`.xlsb` operator path is unchanged. No product VBA, XLAM, ControlId or displayed-wording change was made. Draft PR14 modifies only test tooling.
- The v33 one-run approval is consumed; do not retry it. The new PR4-BASE-10/v34 run requires separate approval. It must start before **2026-11-12 19:50 UTC** to fit the owner's **21:08 UTC** shift stop. No forced Excel close or process kill.

## Evidence and traceability

- v33 stopped at fixture creation with six compile/probe PASS checks, one combined saved-host assertion failure and 196 unrun Settings identities. Both native exits, settings and input preservation verified; delayed audit zero Excel faults. It is not product RED or Settings GREEN.
- The separately approved copied host-state probe returned 0 and recorded `HostObjectAccessible=true`, `HostFullNameMatches=true`, `HostSaved=false`, `OpenWorkbookCount=1`, `MatchingOpenWorkbookCount=1`, `HostFileExists=true`. Excel exited naturally, no delayed Excel fault. Thus the in-memory dirty flag alone failed v33's guard; exact saved host identity persisted.
- Draft PR14 `b8f94241` removes only the after-fixture dirty-flag assertion, retaining exact path, disk file and sole-workbook checks. PR4-BASE-10 manifest `68338D11...` re-pins that one test file; five candidate XLAMs unchanged. The v34 controller `9050B1C4...`, expected file `B7400CF6...` and `codex-settings-policy-v34.md` are in `reports/runtime/settings-policy-v34-pr4-base-10-20261112`. Parser zero errors; checks-only passed 539 exact pins and 202 expected names with no Excel. No v34 approval or dispatch exists.
- Earlier standalone Inventory, Recording and technical B0 replay results remain in handoff 228 and Plan 022; `B0Accepted=false`.

## Do Not Repeat

Do not repeat v32/v32.1/v33, treat the dirty flag as lost saved-workbook identity, weaken path/file/count guards, or combine independent gate totals into 4be-A acceptance. Do not edit the v34 controller to extend its fixed start window.

## Assumptions to Re-verify

Recheck PR heads, clean worker, Excel closure, five package hashes, 539 pins, desktop/input, Excel build and available memory before dispatch. The 202 Settings checks remain unproven on the corrected host.

## Open questions and blockers

One-run v34 approval and packaged GREEN; remaining Event Viewer/detail, guide transfer, other Admin/Operations, retained-formatting, human B0, live-role, restart/binding, full-chain and visible operator acceptance. Residual Execute slowness and the Microsoft-reported Excel exit defect remain separate.

## Immediate next action

Obtain separate approval for PR4-BASE-10 and one v34 run, rerun checks-only at the clean `b8f94241` worker, then dispatch once before 19:50 UTC only if every gate passes.

## Critical references

Architecture v4.11 operator deployment model and D18; Plan 022 PR4-BASE-10 checkpoint; Controls 1.742; draft code PR14; docs PR5; `tests/tooling/Test-Slice4beConfigCommands.ps1`; local v33, host-state probe and v34 folders above.
