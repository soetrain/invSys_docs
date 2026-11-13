# Handoff 226: Worksheet PARTIAL and notice gates on current candidate

Date: 2026-11-11 Pacific. Predecessor: handoff 225.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18-REPLAY-01 and Plan 022. The current five-XLAM candidate now has focused worksheet Upload `PARTIAL` and stale-context notice GREEN; broad Inventory, early Receiving B0 acceptance, 4be-A and Release 1 remain open.

## Current verified state

Last verified 2026-11-11 Pacific: draft code PR11 `work/d18-partial-choice-green` is at `0e014dc8`; draft PR12 `work/d18-worksheet-save-partial` at `c4444f58`; draft PR13 `work/d18-worksheet-notice-fixture` at `f9016668`. Each is stacked on the preceding branch. The original worker checkout `invSys_edit_receipts` is restored clean to `work/d8-inv-edit-receipts` at `550d9ff0`; main code checkout remains at `6d247543` with Justin's modified `AGENTS.md` untouched. Docs branch `work/perplexity-handoff218` has Plan 022 and Controls 1.737 at `3480af9` before this handoff commit. Draft docs PR5 remains open. Unrelated docs deletion, edit and untracked files shown by `git status` were preserved. Excel was closed after v27; recheck before the next test/build.

## Decisions and constraints

- Architecture v4.11 already requires exact `PARTIAL` after inventory application with unconfirmed worksheet result save and requires the fixed tracking-context notice. No new product contract was approved or implemented. PR12 changes only the post-application error outcome. PR13 changes only the packaged test fixture, not operator behavior.
- PR4-BASE-06, 07 and 08 are separate ignored input manifests. BASE08 changes one test helper hash from BASE07 and retains all five BASE06 validation XLAM hashes. The five packages remain validation artifacts, not promoted release packages. Prior one-run dispatches are consumed; never combine their counts or reuse an approval.
- The user requested this session stop within four hours of 2026-11-11 19:18 Pacific, with documentation and this handoff. No Excel process was killed or assisted to close. Continue preserving captured-workbook binding, immutable `System_Key`, unknown user columns, headless authority and unrelated user edits.

## Evidence and traceability

- v23 current packaged behavioral RED: 76/78 exact PASS; only `UploadSaveThrows.ExactCorrelatedOwnerOutcome` and `.ExactOwnerEventReferences` failed. PR12 commit `c4444f58` makes a post-application error `PARTIAL` while keeping report and event references. v24 on BASE06: 78/78 exact packaged GREEN, five compiles, source extraction match, natural exits, restored settings, preserved 539 pins and zero delayed Excel faults. Static maintenance stayed at 391 components, 6,743 procedures, 1,246 scanner candidates, 45 unresolved dynamic calls and 27 oversized-module ratchets without growth. See local `reports/runtime/worksheet-recording-green-v24-pr4-base-06-20261111/` and `worksheet-partial-static-01/`.
- v25 on BASE06: WorksheetResults 64/64, OwnResults 50/50, HeadlessWorksheet 56/56, exact identities and natural exits. WorksheetResults required three recorded OK clicks for the actual fixed tracking notice, so its 64/64 is **attended diagnostic**; the other two runs were unattended. Modal captures are in local `reports/runtime/inventory-results-v25-pr4-base-06-20261111/run-01/`.
- v26 on BASE07 stopped before behavioral checks because the first fixture seam could not find the packaged method. It is a harness setup failure, not product RED. PR13 commit `f9016668` uses the established whole-module seam. v27 on BASE08: WorksheetResults **67/67 exact packaged GREEN**, including the three new notice assertions; five instrumented compiles; no modal or intervention; both native exits verified; settings and 539 pins preserved; delayed audit zero Excel faults. Local evidence: `reports/runtime/inventory-worksheet-notice-v27-pr4-base-08-20261111/`, manifest SHA256 `B5FE3624...`.
- v27 took 26.24 minutes on Excel 16.0.20527.20196, versus v25's attended 41.19 minutes on the same build. Seven `TestInventoryCommands.Execute` calls still totaled 718.2 seconds (maximum 175.0). The different fixture and modal waits confound the comparison; no Office or slowdown fix is claimed. The Microsoft Excel report was already submitted.
- Prior v20 early Receiving B0 replay passed 121/121 technical checks, but its proof explicitly says `B0Accepted=false`. D18-REPLAY-01 requires B0 GREEN before A acceptance and human acceptance remains separate. The current candidate has not yet completed the consolidated A family/control and Release 1 gates.

## Do Not Repeat

Do not rerun v23-v27 or treat v25's attended result, v26's harness failure, or technical B0 replay as acceptance. Do not change the old manifests, promote validation XLAMs, merge draft PRs to main, or revive the historical 50/28 worksheet RED as current-package evidence. Avoid another Excel-exit investigation for this A workstream unless a concrete closure failure recurs.

## Assumptions to Re-verify

Recheck PR heads, baseline hashes, clean worker checkout, Excel closure, desktop input, latest Excel build and unrelated work before any new runtime gate. v27 has no AfterQuit trace; its 2.9-second stdout-quiet-to-Excel-exit interval is not an AfterQuit measurement.

## Open questions and blockers

Retained-formatting named-caller coverage; current-candidate Inventory Recording/Guide selectors and other 4be-A families; early B0 acceptance; live-role, packaged restart/binding, full Release 1 chain, layout/static regressions and visible operator comparison remain unqualified. Existing pre-November passes are on older packages. The underlying Excel exit defect and residual `Execute` slowness remain separate.

## Immediate next action

Prepare a separate BASE08 pinned packaged Recording selector 04 gate (151 identities) against the unchanged five validation XLAMs, check its input hashes and expected identities without Excel, then run one natural-closure attempt under its own authorization.

## Critical references

Architecture v4.11 D18-REPLAY-01/D13 and Inventory worksheet `PARTIAL`; Plan 022 2026-11-11/12 checkpoints; Controls 1.737; `src/Admin/Modules/modAdminInventoryWorksheet.bas`; `tests/tooling/Slice4beInventoryWorksheetResults.ps1`; code PR11/12/13, docs PR5; local v23-v27 evidence folders named above.
