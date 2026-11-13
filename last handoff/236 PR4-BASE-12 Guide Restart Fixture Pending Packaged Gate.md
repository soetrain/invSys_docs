# Handoff 236: Guide restart fixture and pending packaged gate

Date: 2026-11-13 Pacific. Predecessor: handoff 235.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. Current-package guide transfer remains unqualified; PR4-BASE-12 v46 is prepared for one approved packaged run.

## Current verified state

Last verified 2026-11-13: code main `6d247543` retains Justin's unrelated `AGENTS.md` edit. Clean execution worker branch `work/d18-recording-fixture-filter` is `686dbf02`, pushed to soetrain and draft code PR15 (base `work/d18-settings-saved-host`). Docs branch `work/perplexity-handoff218` carries this handoff and Controls 1.748 after commit; unrelated docs edits and untracked files remain untouched. Excel was closed after v45. Five validation XLAMs are unchanged.

## Decisions and constraints

- PR15 changes test fixtures only. No VBA, XLAM, architecture, Plan contract, ControlId or displayed wording changed. Architecture v4.11 permits eligible Settings selection observations during explicit recording; fixtures must account for them.
- Justin authorized closing modals, IDE and programs when needed. Capture any modal before interaction. An intervened run is attended evidence, never unattended GREEN.
- v45's single-run authorization is consumed. PR4-BASE-12 and v46 need separate owner approval. No v46 Excel dispatch has occurred.

## Evidence and traceability

- v45 on PR4-BASE-11 passed the v44 Save Value fixture, then stopped in `Slice4beGuideRestartFixture.ps1` after 11 PASS / one fixture exception / 488 expected identities missing. Both processes exited naturally, settings and pins were preserved, and the delayed Excel audit found zero faults. This is a fixture prerequisite failure, not product behavioral RED or packaged GREEN. Evidence: `reports/runtime/guide-transfer-recorded-v45-pr4-base-11-20261113/`.
- The restart fixture formerly required six journal versions, two actions and four observations. Two Settings selection pairs per save were observed in v44; source predicts fourteen versions, six actions and twelve observations for two saves. That full packaged shape is still **inferred**, not confirmed. The modified fixture writes sanitized structural counts on mismatch. File-only test RED 4 PASS / 1 FAIL, GREEN 5/5; prior recording-filter test still passes 4/4; PS5.1 parsing and diff checks pass.
- PR4-BASE-12 manifest `224F0BE513C207B2666B85618CAFB28A86CAC7CBE0A658EC21543C4F3D39926D` pins 541 files, changing only the restart fixture and adding its test relative to BASE-11. v46 checks-only passed 541 pins and 499 exact expected names with no Excel. Expected file `35D7FDC83902ECFEEBF4F6A73E789ECA86E5D19C7CE24DD93CF707517FB66796`; launcher `0CCFC0114DC0FF2A607F6D7A0726B87B24E5388DF5A318F7B00C01439A09B17D`.

## Do Not Repeat

Do not rerun v45 or call its fixture exception product RED. Do not use the changed fixture under BASE-11 or combine v41-v43 Viewer gates into a full chain. Do not treat the inferred journal shape as observed packaged evidence.

## Assumptions to Re-verify

Recheck owner approval, launcher/manifest/expected hashes, clean worker HEAD, Excel closure, desktop input and current Excel build before v46. The exact packaged journal/guide shape is unresolved. The current PR and remote heads may change.

## Open questions and blockers

The separate PR4-BASE-12/v46 approval is pending. Recorded guide transfer, human Receiving B0 acceptance, broader Admin/Operations coverage, live-role, full-chain, visible operator evidence, 4be-A and Release 1 remain open. Excel visible-start break cause and Execute slowdown remain separate.

## Immediate next action

After owner approval, record its exact input hashes in `v46/session-authorization.md`, recheck checks-only and Excel closure, then run the v46 packaged guide-transfer controller once and classify its structural receipt and exact identities.

## Critical references

Architecture v4.11 D18/D18-REPLAY-01; Plan 022; Controls 1.748; code PR15; docs PR5; `tests/tooling/Slice4beGuideRestartFixture.ps1`; `tests/tooling/Test-Slice4beGuideRestartFixtureShape.ps1`; `reports/runtime/guide-transfer-recorded-v46-pr4-base-12-20261113/codex-guide-transfer-recorded-v46.md`.
