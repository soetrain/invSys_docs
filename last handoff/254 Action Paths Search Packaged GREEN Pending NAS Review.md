# Handoff 254: Action Paths Search packaged GREEN, NAS review pending

Date: 2026-11-20 Pacific. Predecessor: handoff 253.

## Goal and release outcome

Complete 4be-A under Architecture v4.11; approved D18-SEARCH-INPUT-01 now has packaged GREEN, while native NAS usability and operator acceptance remain open.

## Current verified state

Last verified 2026-11-20: code worker `C:\Users\Justin\repos\invSys_edit_receipts`, branch `work/d18-picker-save-repro` at `aa44b245`, pushed to draft [code PR17](https://github.com/soetrain/invSys_fork/pull/17) against `work/d14-packaged-chain-coverage`. Tracked tree is clean; preserve pre-existing untracked `tests/tooling/reports/`. Docs `C:\Users\Justin\repos\invSys_docs`, branch `work/perplexity-handoff218` at `c82afd0` before this handoff, pushed to draft [docs PR8](https://github.com/soetrain/invSys_docs/pull/8) against `work/excel-shutdown-timing`. Preserve unrelated guidance 024 deletion, handoff 067 edit and four untracked user files. Excel count was zero at 18:11 Pacific; Justin had said he could test, but no result was received. No update to `deploy/current` or NAS was made.

## Decisions and constraints

Justin approved D18-SEARCH-INPUT-01. `txtPathSearch` filters only headers loaded by authorized Open/Refresh; typing checks captured session/warehouse locally, clears selected evidence/actions, and does no policy or journal read. Selection/Refresh recheck current policy. Explicit Refresh preserves the same valid selected run; it does not discard an open guide draft merely because headers were reloaded. Controls 1.785 and Plan 022 record the implementation checkpoint. All five final validation XLAMs are in `deploy/validation-action-search-final-20261120/` in the worker. User must load those exact files together for native review; older registered add-ins can mask the change.

## Evidence and traceability

Symptom -> path -> rule: slow native Search -> `frmActionPaths.mSearch_Change` formerly called Core `ListPaths` and reread Config/NAS per key -> Architecture v4.11 D18-SEARCH-INPUT-01. Actual-handler packaged RED on prior five XLAMs: **123 PASS / 4 behavioral FAIL**, no harness failure, natural Excel exit. Final rebuilt five-XLAM candidate: **130/130 GREEN**, five explicit compiles and cold-start reference check, natural Excel exit; report `reports/runtime/slice4be-viewer-published-read/f3256936b0ec44b9a7e4bfe935c61bf2/`. Checks cover local filtering without policy/journal read, selected evidence/action clearing, Refresh/current-policy reselection, cached headers after policy change, target/sign-out invalidation, reader integrity and layout. Static report `reports/runtime/action-search-final-20261120-static/`: 391 components, 6762 procedures, 9 literal/45 unresolved dynamic calls, 193 duplicate-body groups; no unexplained ratchet increase against handoff 251. Package and source reports stay ignored/local. Native keyboard response on the disposable NAS warehouse is not yet proven.

## Do Not Repeat

Do not use or count the disqualified experimental `reports/runtime/4be-action-path-search*-20261120/` builds. Do not replace `deploy/current` on packaged evidence alone. The first 126/1 GREEN attempt used a wrong one-row test expectation; the later added policy check initially reused unchanged Search text and did not fire `Change`; both were corrected before final 130/130.

## Assumptions to Re-verify

Check Excel state, exact loaded XLAM paths, NAS access, candidate hashes, and both PR heads before native or further packaged work. If Justin has already opened Excel, leave it untouched until he completes review.

## Open questions and blockers

Justin's native Search timing/selection result is pending. Same-task How-To/Diagnostic/Compare both for Operations and Admin, an incomplete run, human B0 acceptance, final-candidate Inventory/launcher/Production/full-chain gates and 4be-A acceptance remain open; 4be-B is later.

## Immediate next action

Record Justin's native per-character Search and selection result on the final five-XLAM candidate, including the exact loaded Operations path, before deciding whether the Search increment qualifies for visible NAS usability.

## Critical references

Architecture v4.11 D18-SEARCH-INPUT-01; Plan 022 top checkpoint; Controls 1.785; `src/Operations/Forms/frmActionPaths.frm`; `tests/tooling/Slice4beRecordingReader.ps1`; final validation XLAM folder and ignored reports above; code PR17 and docs PR8.
