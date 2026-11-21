# Handoff 235: Viewer gates and guide recording fixture

Date: 2026-11-13 Pacific. Predecessor: handoff 234.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18 and Plan 022. Current-package Viewer technical gates advanced; recorded guide transfer and overall acceptance remain open.

## Current verified state

Last verified 2026-11-13: code checkout main is 6d247543 with Justin's unrelated AGENTS.md edit. The clean execution worker is on work/d18-recording-fixture-filter at 3a5700c1, published in draft code PR15 (base work/d18-settings-saved-host at b8f94241). Docs branch work/perplexity-handoff218 was at c30f4fe before this handoff; unrelated docs edits and untracked files were preserved. Excel is closed. Five validation XLAM hashes are unchanged. Controls 1.747 and Plan 022 record today's checkpoint.

## Decisions and constraints

- No product VBA, XLAM, architecture, plan contract, ControlId or displayed wording changed. PR15 is test tooling only.
- Justin authorized closing modals, the IDE and programs when needed. Capture any modal first. An intervened run is attended diagnostic evidence, never unattended GREEN.
- v41-v44 one-run authorizations are consumed. PR4-BASE-11 and v45 need the pending explicit owner approval; no v45 Excel dispatch has occurred. The v45 launcher has a 2026-11-13 20:45 UTC start cutoff. The shift ends by 21:45 UTC.

## Evidence and traceability

- PR4-BASE-10 v41 hidden-start Viewer published read/recording: 80/80 exact, native Excel exit, settings and 539 pins preserved, no intervention, delayed audit zero faults. This supports initial host visibility as the v40 break condition but does not prove the underlying Excel cause.
- Separate v42 groups/publication: 82/82 exact. Separate v43 refresh failure: 16/16 exact. Both had natural exits, restored settings, preserved pins, no intervention and zero delayed Excel faults. Evidence is under reports/runtime/viewer-groups-publication-v42-pr4-base-10-20261113 and viewer-refresh-v43-pr4-base-10-20261113.
- v44 recorded guide transfer stopped at 11 PASS, one fixture exception and 488 missing checks. The sanitized activity count showed two ADMIN_SETTINGS_SELECT_CONFIG request/selection pairs alongside one ADMIN_SETTINGS_SAVE_VALUE request/completion pair. Source path: tests/tooling/Slice4beRecordingFixture.ps1, SaveRecordedSetting. This is a fixture prerequisite failure, not product behavioral RED. Both native exits, restoration and pin preservation passed; delayed audit found zero faults.
- PR15 changes the fixture to select the exact Save Value pair while preserving unrelated selection records. File-only test: RED 3 PASS / 1 FAIL (ExtraSelections), GREEN 4/4, PowerShell parse clean. PR4-BASE-11 manifest SHA256 23EEDA5107EDCC6C2739B64DC149DAFD8C7911C1148A2CB15B22A97410816336 has 540 pins: one changed fixture, one new test, all other old pins and five XLAMs unchanged. v45 checks-only passed 540 pins and 499 historical exact identities; launcher SHA256 619287524BCD48E120DCA462AD7A877A44518E85C8268EB744B798BB4E865C37, expected file SHA256 50DEB79F572219EE0D6AD139457B51BC2CB68542B5AB7AA0A89EE790A550E599.

## Do Not Repeat

Do not rerun v40 or v44, call the v44 exception product RED, combine v41-v43 as one chain, or use the new fixture under the old BASE-10 manifest. The 499-check transfer gate has not passed on the current packages.

## Assumptions to Re-verify

Recheck owner approval, cutoff, Excel closure, worker and docs branch cleanliness, package hashes, desktop input and current Excel build before v45. Whether anyone interacted with v40's VBA break remains unconfirmed.

## Open questions and blockers

The PR4-BASE-11 re-pin and one v45 packaged run await owner approval. Recorded guide transfer, human Receiving B0 acceptance, broader Admin/Operations coverage, live-role, full-chain, visible operator evidence, 4be-A and Release 1 remain open. The Excel visible-start break cause and Execute slowdown remain separate.

## Immediate next action

If the owner approves within the cutoff, record the exact approval hashes in v45 session-authorization.md and run the prepared v45 guide-transfer controller once; classify its exact identities and natural closure before any product change.

## Critical references

Architecture v4.11 D18 and D18-REPLAY-01; Plan 022; Controls 1.747; code PR15; docs PR5; worker tests/tooling/Slice4beRecordingFixture.ps1 and Test-Slice4beRecordingFixtureFiltering.ps1; reports/runtime/pr4-base-11-recording-fixture-20261113/frozen-input-pins-pr4-base-11-3a5700c1.json; reports/runtime/guide-transfer-recorded-v45-pr4-base-11-20261113/Invoke-GuideTransferRecordedV45.ps1.
