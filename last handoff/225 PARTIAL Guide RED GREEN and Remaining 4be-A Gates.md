# Handoff 225: PARTIAL guide RED/GREEN and remaining 4be-A gates

Date: 2026-11-11 Pacific. Predecessor: handoff 224.

## Goal and release outcome

Complete invSys Release 1 Slice 4be-A under Architecture v4.11 and Plan 022. The existing Inventory Create/Upload PARTIAL guide contract now has focused packaged RED/GREEN evidence; broad Inventory GREEN, 4be-A and Release 1 acceptance remain open.

## Current verified state

Last verified 2026-11-11 Pacific: code PR4 working branch `work/d8-inv-edit-receipts` is clean at `550d9ff08fcce2d78c1a059a51ca7065928c9c66`; draft PR11 `work/d18-partial-choice-green` is clean and published to soetrain at `0e014dc8d22b64744288ddd58de8466076422786`, stacked on PR4. Docs branch `work/perplexity-handoff218` was at evidence commit `505a7a532b24a447674a167d7d5ead8780bfed0f` before this handoff commit; it is the head of draft docs PR5 on the justinwj fork. Plan 022 and Controls 1.735 record the focused evidence. Main code checkout keeps Justin's modified AGENTS.md; docs checkout keeps the unrelated changes listed by `git status`. Preserve them. Excel was closed after v22 and after the static scan; recheck before any new build/test.

## Decisions and constraints

- Architecture v4.11 already defines PARTIAL for Inventory Create/Upload; no architecture amendment was needed. Commit `0e014dc8` adds only `PARTIAL` to `modExpectationDraft.Choices`; the existing `modActivityCatalog.Outcome` filter still governs each control.
- Justin separately approved PR4-BASE-04/v21 RED and PR4-BASE-05/v22 GREEN. Both were one-run authorizations and are consumed. Do not retry or combine their counts. Five candidate XLAMs for v22 are validation packages, not promoted packages.
- Perplexity Search could not write to soetrain; Codex committed/pushed PR11 there. Docs PR5's head is on justinwj and was advanced without changing its soetrain base. Future cross-account work must check remote heads before pushing; do not force-push or merge to main.

## Evidence and traceability

- v19 on PR4-BASE-03: Recording.Uom 68/68 exact, natural closure. v20 early Receiving replay: 121/121 technical checks; `B0Accepted=false` in its proof. Neither is combined with earlier series.
- v21 on unchanged old packages and test-only commit `d5f0170b`: 80 exact checks, 76 PASS and exactly four Create/Upload PARTIAL choice FAIL; two UOM checks and 74 retained guide checks PASS. Natural closure, restored settings, preserved inputs, zero delayed Excel faults. Evidence: `reports/runtime/partial-choices-red-v21-pr4-base-04-20261111/`.
- v22 on PR4-BASE-05 manifest SHA256 `2280127D...` and rebuilt five XLAMs: 80/80 exact PASS including the prior four failures, five instrumented package compiles, retained guide layout checks, worker exit 0, both native exits verified, stdout-quiet to Excel exit 1.57 seconds, settings restored, 539 pins preserved, no termination, zero delayed Excel faults. The five XLAMs matched source extraction. Evidence: `reports/runtime/partial-choices-green-v22-pr4-base-05-20261111/`.
- Static maintenance regenerated from the green tree at `reports/runtime/partial-choices-green-static-01/`: 391 components, 6,743 procedures, 1,246 scanner candidates and 27 oversized-module ratchets. Scanner warnings include existing retired-ROW and unresolved dynamic-call findings; this is a generated report, not global static acceptance.

## Do Not Repeat

Do not rerun v21 or v22, normalize frozen tooling line endings, replace PR4-BASE-01 through BASE-04 manifests, or treat technical B0 replay as B0 acceptance. Do not promote PR11 or PR4 to main on this evidence alone. The Microsoft Excel exit report is already submitted; the underlying exit defect and intermittent Execute slowdown are still separate open issues.

## Assumptions to Re-verify

Recheck PR and branch heads, Excel closure, desktop input, candidate hashes, and untouched unrelated work before another runtime run. The v22 controller found no AfterQuit trace, so its 1.57-second stdout-quiet interval is the available exit timing, not an AfterQuit measurement.

## Open questions and blockers

The 28 worksheet-recording failures are historical RED on the older `validation-inventory-uom-02` package set; the current PR4/PR11 packages have different hashes and have not run that selector. The sanitized evidence is in `expert guidance docs/Worksheet Recording RED Evidence for Review.md`. An ignored v23 controller and 78-name expected file are prepared under `reports/runtime/worksheet-recording-evidence-v23-pr4-base-05-20261111/`; checks-only passed with no Excel. Its launcher SHA256 is B16136D6..., expected checks D3AB5977..., and owner approval/dispatch are absent. The worker branch was restored clean at 550d9ff0. Remaining Inventory selectors and retained formatting callers, relevant source/static regression review, live-role, restart, full Release 1 chain and visible operator evidence are open. The early Receiving B0 proof is technically GREEN but explicitly not accepted. 4be-A is incomplete.

## Immediate next action

After a separate owner approval naming v23 and its frozen hashes, recheck its checks-only gate and run the single prepared worksheet `Evidence` selector on PR4-BASE-05; map only current qualified failures to Architecture v4.11 before changing product source.

## Critical references

Architecture v4.11 Inventory PARTIAL and worksheet observation contracts; Plan 022 2026-11-11 checkpoint; Controls 1.735; `src/Core/Modules/modExpectationDraft.bas`; `tests/tooling/Slice4beInventoryPartialChoices.ps1`, `Slice4beInventoryWorksheetRecording.ps1`; `expert guidance docs/Worksheet Recording RED Evidence for Review.md`; draft code PR4 and PR11; draft docs PR5; local v21/v22/green-static and v23 preparation folders above.
