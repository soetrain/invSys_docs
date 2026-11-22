# Handoff 260: D18 local guide Search candidate; NAS review open

Date: 2026-11-22 Pacific. Predecessor: handoff 259.

## Goal and release outcome

Complete Slice 4be-A under Architecture v4.11 and Plan 022. This increment implements approved D18-SEARCH-INPUT-02 for Choose tracked actions and Published guides; visible NAS acceptance and broader 4be-A remain open.

## Current verified state

Last verified 2026-11-22: code worker `C:\Users\Justin\repos\invSys_edit_receipts` is `work/d18-picker-save-repro` at `a062655a`, published to draft [PR17](https://github.com/soetrain/invSys_fork/pull/17) targeting `work/d14-packaged-chain-coverage`. Docs `C:\Users\Justin\repos\invSys_docs` is `work/perplexity-handoff218` at `fa7e7e1`, published to draft [PR8](https://github.com/soetrain/invSys_docs/pull/8) head `work/d18-picker-save-repro-docs`, base `work/excel-shutdown-timing`. Both PR heads and draft states were verified. Code has only pre-existing untracked `tests/tooling/reports/`; unrelated docs deletion, edits and untracked files remain untouched. Excel is closed. Isolated candidate `C:\Users\Justin\repos\invSys_edit_receipts\deploy\validation-search-extension-20261122` has not been promoted to `deploy/current` or NAS.

## Decisions and constraints

Justin approved D18-SEARCH-INPUT-02 in this conversation. Architecture v4.11, Plan 022 and Controls 1.797 are synchronized. Search changes filter a Core-authorized metadata snapshot in Operations with an in-memory exact context check. Picker selection and Create, guide selection and Refresh retain Core guards. No guide schema, ControlId or recording rule changed. D18-RECORDING-LIST-01 remains proposed, not approved. No PIN or service credential is stored here.

## Evidence and traceability

- Behavioral RED on old packages: Published guides 178 PASS / two new selection failures, `reports/runtime/slice4be-viewer-published-read/c02624a6d993450f9a49731da98f0472/red.json`. Old-package native run: 110 PASS / nine FAIL, including first and subsequent key/selection failures; form unavailable at no-match, so no complete quantitative baseline. Old picker native baseline also lost the form at no-match and is not a timing qualification.
- Corrected packaged handlers: picker 75/75 (including hidden plus new exact ActivityIds and guarded Create), reader/expectation 182/182. Native keyboard-to-painted-list: picker 81/81 with three action rows, first key 64.1 ms, later keys 8.8–10.9 ms; Published guides 194/194 with two guide versions, first key 73.9 ms, later keys 7.4–7.6 ms. Limits were 1500 ms first key and 750 ms later keys. Native files: `reports/runtime/slice4be-viewer-published-read/52cd19cf22b74910b990b18e180000ed/native-frmGuideActionPicker-search.json` and `.../e39ae4c1f7d44adc98dbfb5748fb3904/native-frmActionPathLibrary-search.json`.
- Test-only Core counters showed zero per-key guide-library/policy reads and Core revalidation on Create/Refresh. Reader policy, exact-version chain, source preservation, layout, Close button and native X passed. Retained transfer passed 21/21; `reports/runtime/guide-transfer-controller/fcaf70d7c9aa42ae9198bd5d66bfeafb/closure.json` confirms settings restored, packages preserved and Excel closed. No Excel Application Error 1000/1002 occurred in the candidate-run window after 02:00 Pacific.
- Five-XLAM candidate hashes: Admin `468BE5502AA0B836140E2D858E7A397300E901BD98821EB31B9B8F5BC5D70F69`, Core `8F936BB2B958F9927A13B59F78F9D11D6305CFF8117F933783F60FE606C626B5`, Designs Domain `C6E43B4992152901DACA3950372E4B1AF8EB55F54B8093B8C0045E930B134B51`, Inventory Domain `0719151D2724BA941E886F8BFF0222E44605761A6C2B8DE62FAB7994E88E7FF3`, Operations `2505EBE5C286E08F46BD994F493CC9B9A0763BE3E8BB58D747E7F6245458D331`. All five compiled in the packaged runs.
- Static report `reports/runtime/action-search-input02-static-20261122/`: 392 components, 6,772 procedures, 1,265 candidates, 1,060 dynamic roots and 30 warnings. Relative to prior Close checkpoint, two local-filter procedures were added; candidates, dynamic roots and warnings did not grow.

## Do Not Repeat

The first picker RED attempt stopped at an unrelated Settings fixture exception; it was not behavioral RED. An early candidate run hit a test adapter Null conversion after product cleared selection; that run was attended and invalid as GREEN, and the adapter was corrected. The old native baselines did not complete no-match after their forms became unavailable; do not count them as quantitative comparisons.

## Assumptions to Re-verify

Recheck Excel closure and exact candidate hashes before loading it. Native fixtures had only three actions and two guide versions; responsiveness on the larger disposable NAS list is not yet proven. The prior signed-in review session was closed for packaged tests; Justin must sign in privately again.

## Open questions and blockers

Justin's visible review of both searches and Published guides Close/X on the exact candidate is pending. A child runner executing while Search invalidates selection has not been separately exercised. Slice 4be-A, B0 and Release 1 acceptance remain open.

## Immediate next action

Load the five candidate XLAMs into a private Excel review session for disposable WHT7025AE, then obtain Justin's visible Search actions, Published guides Search and Close/X results before any acceptance claim.

## Critical references

Architecture v4.11 D18-SEARCH-INPUT-02; Plan 022; Controls 1.797; `frmGuideActionPicker.mSearch_Change`, `frmActionPathLibrary.mSearch_Change`, `modCuratedGuideSource.ReadRows`, `Slice4beGuideSearchNative.ps1`, `Slice4beGuideActionCuration.ps1`, `Slice4beGuideLibrary.ps1`.
