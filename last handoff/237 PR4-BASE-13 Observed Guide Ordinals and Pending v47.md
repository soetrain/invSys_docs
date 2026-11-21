# Handoff 237: Observed guide ordinals and pending v47

Date: 2026-11-13 Pacific. Predecessor: handoff 236.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. Guide transfer remains unqualified; v47 is prepared to test the packaged-observed Settings action order.

## Current verified state

Last verified 2026-11-13: code main `6d247543` retains Justin's unrelated `AGENTS.md` edit. Clean execution worker `work/d18-recording-fixture-filter` is `d6fb040c`, pushed to soetrain in draft code PR15 (base `work/d18-settings-saved-host`). Docs branch `work/perplexity-handoff218` carries Controls 1.749 and this handoff after commit. Unrelated docs changes and untracked files remain untouched. Excel closed naturally after v46. Five validation XLAMs are unchanged.

## Decisions and constraints

- PR15 changes test tooling only; no VBA, XLAM, architecture, Plan contract, ControlId or displayed wording changed. v46's one-run approval is consumed. PR4-BASE-13/v47 requires separate owner approval; no v47 Excel dispatch has occurred.
- Capture any modal before interaction. A run requiring intervention is attended evidence, not an unattended GREEN gate. Do not retry a consumed packaged run or assist Excel teardown.

## Evidence and traceability

- v46 on PR4-BASE-12 passed 11 checks, then raised one restart-fixture exception and left 488 identities unrun. Sanitized `guide-restart-journal-shape.json` recorded 14 journal versions, valid chain, Lifecycle Stopped, six actions, 12 observations, four `ADMIN_SETTINGS_SELECT_CONFIG` request/selection pairs and two `ADMIN_SETTINGS_SAVE_VALUE` request/completion pairs. Save Value ordinals were **3 and 6**, while the fixture expected 2 and 5. Both processes exited naturally; settings/pins were preserved; delayed audit found zero Excel faults. This is a fixture prerequisite failure, not product RED. Evidence: `reports/runtime/guide-transfer-recorded-v46-pr4-base-12-20261113/`.
- Updating the file-only ordinal test to 3/6 gave focused RED 4 PASS / 1 FAIL; changing the fixture gave GREEN 5/5. The prior recording-filter test remains GREEN 4/4; PowerShell 5.1 parsing and diff checks passed. Commit `d6fb040c` is on draft code PR15.
- Separate PR4-BASE-13 manifest SHA256 `B3DCF39A6DFB52B5312FB98C592FFBD04F8248B795C868BF4E5440A759D57C12` pins 541 files, differing from BASE-12 only in those two hashes. v47 checks-only passed 541 pins and 499 exact names with no Excel. Expected file SHA256 `17AADDDE7F27F477AF48489996520D3A3100058C3A555303219FC216B1198D1D`; launcher SHA256 `DA224837C94E2291E24564D0E5BCB25A8F86FA90ABB723A6DCCC59F5EDB8C5A3`.

## Do Not Repeat

Do not rerun v45 or v46, count either fixture exception as product RED, reuse their consumed authorizations, or use the changed fixture under BASE-12. Do not treat the v47 checks-only pass as packaged GREEN.

## Assumptions to Re-verify

Recheck owner approval, three hashes, clean worker HEAD, Excel closure, desktop input and current Excel build before v47. Later guide assertions after the journal gate are still unproven on current packages.

## Open questions and blockers

PR4-BASE-13/v47 owner approval is pending. Recorded guide transfer, human Receiving B0 acceptance, broader Admin/Operations, live-role, full-chain, visible operator evidence, 4be-A and Release 1 remain open. Excel visible-start break cause and Execute slowdown remain separate.

## Immediate next action

If approved before the fixed start cutoff, record the exact hashes in `v47/session-authorization.md`, recheck dispatch gates and run the packaged v47 controller once; classify exact identities and natural closure.

## Critical references

Architecture v4.11 D18/D18-REPLAY-01; Plan 022; Controls 1.749; code PR15; docs PR5; `tests/tooling/Slice4beGuideRestartFixture.ps1`; `tests/tooling/Test-Slice4beGuideRestartFixtureShape.ps1`; `reports/runtime/guide-transfer-recorded-v47-pr4-base-13-20261113/codex-guide-transfer-recorded-v47.md`.
