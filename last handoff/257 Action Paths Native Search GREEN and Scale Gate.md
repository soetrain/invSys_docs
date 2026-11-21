# Handoff 257: Action Paths native Search GREEN and scale gate

Date: 2026-11-20 Pacific. Predecessor: handoff 256.

## Goal and release outcome

Complete Slice 4be-A under Architecture v4.11. Approved D18-SEARCH-INPUT-01 now has focused packaged native-keyboard evidence; larger-library NAS usability and final acceptance remain open.

## Current verified state

Last verified 2026-11-20: code worker `C:\Users\Justin\repos\invSys_edit_receipts`, `work/d18-picker-save-repro` at `660473df`, pushed to draft [PR17](https://github.com/soetrain/invSys_fork/pull/17) targeting `work/d14-packaged-chain-coverage`. Docs `work/perplexity-handoff218` at `8e716c1`, pushed to draft [PR8](https://github.com/soetrain/invSys_docs/pull/8) targeting `work/excel-shutdown-timing`. Excel is closed. Only the pre-existing code `tests/tooling/reports/` and unrelated docs changes remain untracked/dirty; `invSys_fork/main` has an unrelated AGENTS.md edit. No product source, release XLAM, `deploy/current` or NAS file changed.

## Decisions and constraints

D18-SEARCH-INPUT-01 remains approved and unchanged. The saved disposable probe copies and saved operator host are test setup. The native fixture uses one validated process-owned click to focus Search, then continuous typing; a repeated click altered the text sequence. Preserve the prior Excel startup-entry snapshot at ignored `reports/runtime/action-search-startup-before-20261120.json` and restore those two entries when candidate review ends. Do not infer the earlier slow session's loaded XLAM from startup registration alone.

## Evidence and traceability

The focused visible packaged fixture passed **152/152**, including six native Search steps with seven disposable recordings. Exact text, row count and no selected row passed; key-to-result times were 48.6, 6.1, 5.6, 5.7, 49.7 and 4.5 ms. All five disposable XLAMs compiled. Excel exited naturally, the saved host's identity and bytes were preserved, and the candidate XLAM hashes were unchanged. Ignored local result: `reports/runtime/slice4be-viewer-published-read/717cc49005144623916f65c9630c3867/`; timing JSON SHA-256 `21EBBDCF7E6EAE857DA7AE373C734043A177AAC5951CE71616E0494467EC9AC1`. Plan 022, Controls 1.789 and report 026 record the limit: this is not a pinned final candidate, larger NAS library, 4be-A or Release 1 acceptance.

Symptom -> fixture cause -> requirement -> proof: native Search initially stopped before input because unsaved instrumented packages failed Viewer caption assignment; then missed native textbox focus; then a fixture incorrectly expected one row for `Rec`, though all seven titles matched; then repeated clicks replaced a later character. Saved copies, saved host, one validated click, and exact fixture row expectations yielded 152/152 under D18-SEARCH-INPUT-01.

## Do Not Repeat

Do not classify any earlier setup failure as behavioral Search RED. Do not click Search before every key; one initial click supported the continuous sequence. Do not promote the candidate or claim large-library usability from seven rows.

## Assumptions to Re-verify

Recheck actual loaded XLAM paths/hashes, Excel startup entries, display/session state, current Excel build, PR heads and candidate hashes before another native run. The historical slow session's loaded Operations XLAM is still unknown.

## Open questions and blockers

Larger-library NAS Search responsiveness, final candidate qualification, same-task How-To/Diagnostic/Compare, B0, full-chain and visible 4be-A acceptance remain open. No new architecture approval is pending for Search.

## Immediate next action

Prepare a controlled larger-recording fixture and pinned expected identities on the unchanged candidate, then test native Search through the packaged form before visible NAS operator review.

## Critical references

Architecture v4.11 D18-SEARCH-INPUT-01; Plan 022; Controls 1.789; report 026; `frmActionPaths.mSearch_Change`; `Slice4beRecordingReader.ps1`, `Slice4beRecordingSearchNative.ps1`, `Test-Slice4beConfigCommands.ps1`; code PR17 and docs PR8.
