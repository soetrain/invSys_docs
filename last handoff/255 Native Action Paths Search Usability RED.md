# Handoff 255: native Action Paths Search usability RED

Date: 2026-11-20 Pacific. Predecessor: handoff 254.

## Goal and release outcome

Complete 4be-A under Architecture v4.11; Action Paths Search remains a native usability blocker despite packaged GREEN.

## Current verified state

Last verified 2026-11-20: code worker `work/d18-picker-save-repro` at `aa44b245`, draft [code PR17](https://github.com/soetrain/invSys_fork/pull/17), with no tracked changes and pre-existing untracked `tests/tooling/reports/`. Docs `work/perplexity-handoff218` at `96e8ad1` before this handoff, draft [docs PR8](https://github.com/soetrain/invSys_docs/pull/8). Preserve unrelated guidance 024 deletion, handoff 067 edit and four untracked user files. Excel was closed at the read-only check. The isolated candidate is still in `deploy/validation-action-search-final-20261120/`; `deploy/current` and NAS were not changed.

## Decisions and constraints

Approved D18-SEARCH-INPUT-01 remains operative. No contract or product change was made in this session. Controls 1.786 and Plan 022 now mark native usability RED. The expert-facing, sanitized report is `expert guidance docs/026 Action Paths Search Native Performance Problem Report.md`. D13 does not apply to this documentation-only evidence correction. Do not equate packaged GREEN with native acceptance.

## Evidence and traceability

Justin reports extremely slow typing in Action Paths Search; Viewer Events Search and the existing picker are responsive. Prior actual-handler packaged RED was 123 PASS / 4 behavioral FAIL; the isolated candidate reached 130/130 packaged GREEN and five compiles. Those tests set TextBox.Value through the real handler but do not measure native keyboard/focus/paint or the NAS session. Exact XLAM paths loaded during Justin's slow review were not verified. See report 026 for source path, distinct package hashes and bounded expert questions. Symptom -> suspected boundary remains unresolved -> Architecture v4.11 D18-SEARCH-INPUT-01 native responsiveness requirement -> required native packaged RED and visible NAS retest.

## Do Not Repeat

Do not attribute the delay to NAS, MSForms redraw, RDP or stale XLAMs without same-session measurements. Do not promote the candidate or count 4be-A accepted from 130/130. Keep earlier disqualified experimental builds excluded.

## Assumptions to Re-verify

Recheck exact loaded XLAM paths/hashes, Excel state, session/display mode, both PR heads and code/docs worktree status before any native measurement or fix.

## Open questions and blockers

The slow keypress boundary and exact loaded package are unknown. Expert guidance and native timing are pending. Same-task How-To/Diagnostic/Compare both, incomplete-run, human B0 and final Release 1 acceptance gates remain open.

## Immediate next action

Verify the loaded Operations XLAM and capture redacted per-stage timings for one native Search keypress against Events Search and the working picker in the same operator session before changing implementation.

## Critical references

Architecture v4.11 D18-SEARCH-INPUT-01; Plan 022 top checkpoint; Controls 1.786; report 026; `frmActionPaths.mSearch_Change`, `FilterLoadedPaths`, `ClearSelectedRecording`; `frmInventoryViewer.mTxtSearch_Change`; `frmGuideActionPicker.mSearch_Change`; code PR17; docs PR8.
