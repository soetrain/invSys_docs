# Handoff 227: PR4-BASE-08 A-family gates and Settings memory failure

Date: 2026-11-12 Pacific. Predecessor: handoff 226.

## Goal and release outcome

Complete Release 1 Slice 4be-A under Architecture v4.11 D18-REPLAY-01 and Plan 022. Current-candidate Inventory Recording/Guide, Admin Tester/Create recording and technical Receiving replay have qualified; Settings, broad 4be-A and Release 1 acceptance remain open.

## Current verified state

Last verified 2026-11-12 Pacific: code main is `6d247543` with Justin's `AGENTS.md` edit untouched. The worker checkout is restored clean on `work/d8-inv-edit-receipts` at `550d9ff0`. Validation source is draft code PR13 `f9016668`, stacked on PR12 `c4444f58` and PR11 `0e014dc8`; the five candidate XLAMs and 539 PR4-BASE-08 pins were unchanged. Docs branch `work/perplexity-handoff218` contains the Plan 022 / Controls 1.738 checkpoint at `f8c652f`; this handoff is the next docs commit. Unrelated docs deletion, edit and untracked files remain untouched. Excel and Perplexity are closed; recheck before runtime.

## Decisions and constraints

- The current Architecture v4.11 contract is unchanged; no new product behavior, ControlId, XLAM or source implementation was made. Plan 022 and Controls 1.738 record the same acceptance limits.
- v32 was a one-use Settings dispatch. Its attended, memory-constrained result is neither meaningful product RED nor Settings GREEN. Do not retry it or reuse its authorization.
- A distinct v32.1 controller is prepared in ignored runtime files. Checks-only passed, with no Excel launched. Its approval file is absent and dispatch has not occurred. A fresh one-run owner authorization is required.
- Preserve captured-workbook binding, exact immutable `System_Key`, unknown columns, headless Core/Domain authority, unrelated user edits, and natural Excel cleanup. Do not promote validation packages or combine independent gate totals into 4be-A acceptance.

## Evidence and traceability

- `inventory-recording-v28-pr4-base-08-20261112`: selector 04 151/151 exact. v29 run 05 65/65 exact; its run 06 worker printed 75 PASS but observer file sharing prevented qualification. Separate v29.1 runs 06-09 passed 75/75, 68/68, 80/80 and 73/73 exact. v30 Admin Tester/Create passed 163/163 and 170/170 exact. Every qualified run exited naturally, restored settings and pins, and had zero delayed Excel faults.
- v31 checks-only found a controller initializer error before dispatch. Corrected v31.1 passed 121/121 exact early Receiving replay checks, including ten proof-critical identities and four fresh scope flags, with natural closure and zero delayed faults. Its proof says `B0Accepted=false`; human B0 acceptance is open.
- `settings-policy-v32-pr4-base-08-20261112`: Excel displayed captured insufficient-memory dialogs at 82.6/85.7 GiB system commit while Perplexity models used about 44.5 GiB. After the app closed, commit pressure fell and Excel resumed. The attended worker ended 134 PASS, one `Harness.Exception.packaged Settings tracking surface: FAIL` at `SelectTarget` status 6 (`WH_CONFIG_INVALID`), and 68 missing expected identities. Source: `tests/tooling/Test-Slice4beConfigCommands.ps1:1547` -> Core `modNasConnection.SelectWarehouseTargetForAutomation`; environmental causation is not proven. Both native exits verified, settings/pins preserved, no termination, delayed audit zero faults. Modal captures and closure are in that local folder.
- Static maintenance on current source: six packages, 391 components, 6,743 procedures, 1,043 dynamic roots, 1,246 candidates, 27 oversized ratchets; candidate IDs and ratchets match the retained baseline. Three Config source tests passed 13/13, 12/12 and 33/33 with explicit RepoRoot. These do not substitute for packaged acceptance.
- Prepared `settings-policy-v32-1-pr4-base-08-20261112/Invoke-SettingsPolicyV321.ps1`: SHA256 `BC2296EABF3E7BE0DD0FFB8C008034B38327247FA16CE5ABA9F3BA8FC302AE23`; the unchanged expected-identity file hashes `19DD2A45D8033B4466073B2D9FBB74729CB764B42CC403A4D6AF6B10EC11C3BC`. Parser 0 errors; checks-only verified 539 pins and 202 identities. It requires a new `session-authorization.md` naming PR4-BASE-08, v32.1, manifest `B5FE3624F66BA7F59C0711683BB38F149C543B5AEB94AF2B2D1D304C589F4534`, expected hash and launcher hash. Start cutoff is 2026-11-12 19:35 UTC; after that, prepare a new bounded window instead of editing this launcher.

## Do Not Repeat

Do not rerun v29, v31 or v32, infer a product defect from their harness/environment failures, or claim technical replay as human B0 acceptance. Do not restart the Excel-exit investigation without a concrete new closure failure. The original v32 memory modal was captured before any acknowledgment; its run is attended.

## Assumptions to Re-verify

Before a new runtime gate, recheck the five package hashes, 539 pins, Excel closure, system commit headroom, desktop input, branch positions, and any new Excel build. PR11-13 and docs PR5 were draft at the last check; remote state can change.

## Open questions and blockers

Settings 202-check current-candidate qualification; remaining Event Viewer/detail, guide transfer, other Admin and Operations families, retained-formatting callers, live-role, packaged restart/binding, full Release 1 chain, layout and visible operator comparison; human B0 and broad 4be-A acceptance. Residual `Execute` slowness and the Microsoft-reported Excel exit defect remain separate.

## Immediate next action

Obtain one-run approval for prepared v32.1, recheck its checks-only, package and memory gates, then run the 202-check visible packaged Settings route once with natural closure.

## Critical references

Architecture v4.11 D18-REPLAY-01/D13; Plan 022 PR4-BASE-08 checkpoint; Controls 1.738; `tests/tooling/Test-Slice4beConfigCommands.ps1`; code PR11-13; docs PR5; local v28-v32 evidence and v32.1 controller folders named above.