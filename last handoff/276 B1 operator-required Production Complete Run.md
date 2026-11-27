# 276 B1 operator-required Production Complete Run

## Goal and release outcome

Qualify `PRODUCTION_RUN_COMPLETE` under D18-REPLAY-01 as an operator-required guide step. Production owns the deliberate click, authorization, Check In prerequisites, persistence and outcome. Core owns replay attribution and verification.

## Current verified state

Last verified November 27, 2026. The code and test head is `431c76472f5eee35e9d9d8fa730f252a9839cc21` on `work/d18-picker-save-repro`. The implementation commits are `83e057e68c79b737fe275f312a2528ca89189639` and `d1c2dc0fa945cdcd0215712b13fc8234ea646102`. The Architecture, Controls and Plan 022 changes are in `7cf2cd34d61676209357f2bce10d681bcb8b1c4a` on `work/perplexity-handoff218`, published as `work/d18-picker-save-repro-docs`. The starting code/docs pair was `d8e5ad0e` / `28a71185`. Existing unrelated changes in the main checkout, tooling reports and docs working tree belong to other work.

The isolated catalog 48 build source has Git `src` tree `6240e2146d2412ba45e79d2e3def676c422fd316`. To reproduce that source, start at `d1c2dc0f` and take just these six archived files from `27cb6d5b`. The feature branch retains its existing catalog 49 work.

- `src/Admin/Forms/frmAddInventoryItem.frm` and `src/Admin/Modules/modAdmin.bas`
- `src/Core/Modules/modActivityCatalog.bas`, `modActivityReferences.bas`, `modInventoryActivityCodes.bas` and `modInventoryRetirementCommands.bas`

Isolated qualification input: `deploy/validation-b1-production-operator-04` under the active code worktree. Operations cold startup and all five pristine package compiles passed.

| Package | Build identity | Candidate SHA-256 |
|---|---|---|
| invSys.Admin.xlam | `5dfe137bc3974ab9bd32575460e7fdce` | `30A7C15872FA00CA9646216FEB3FB7AC29D6B642A995840781E1ACEB3FE27448` |
| invSys.Core.xlam | `79de29a3387648b688d7971c2167a041` | `24C4287795DF3A498C0F2CA464F6A936F86185C67932039672793A3DAC9DB8AC` |
| invSys.Designs.Domain.xlam | `00ad328506d346d98e8338fcf8c73035` | `2D52403C5A21276FB6B0526E4211A2B05132F96FEECF7DAA2A8E64E181F59D62` |
| invSys.Inventory.Domain.xlam | `6d0c83f855a6448d962047d2bf408e53` | `A659D3D1A52BE711F71A0ECABC4D10F9E1FCE9EBC78D922479E03324EF077988` |
| invSys.Operations.xlam | `48c9bbf282da40f991418e4c54878ce2` | `5808277DA5FB0DB5986F8625126CF371932EC63EBB193E4979FE47138CA58411` |

## Decisions and constraints

1. Justin approved honoring an explicit worksheet output across ordinary Production, including with replay closed. The Process filter supplies a default only when there is no selected output. This fixed an observed debit/creation against the first output when the operator had selected the second.
2. The exact guide/profile identifies the required action. At eligibility, a v2 local ExecutionRun appends a stable StepId and fresh BindingId. The operator selects in Production and clicks **Bind selected Production output** in Run How-To. The selection belongs to the captured form, workbook, actor/session, destination and recording. The reusable identity includes its checked-in batch/node generation and released recipe state. A worksheet draft has a unique ticket in an owned staging column. Both branches fingerprint the prepared output, entered quantity and allocations. Version 1 remains readable for existing runs.
3. The ordinary **Complete Run** handler links one fresh REQUESTED ActivityId to the binding. It checks the actual selected state again before invoking the completion owner. **Check completion evidence** requires the same sequence, activity and CONFIRMED outcome, plus every submitted Inventory event Applied in the loaded publication. Refresh Events and check again when the publication predates completion. The owner call-lifetime guard is released independently of the return notification. A later step requires Next after positive proof.
4. Stop or Close revokes the live selection. Reopening preserves immutable history. A new setup starts a new run. Stale selections and other ordinary successes receive no credit. Whole-guide Verify checks credited identities. Ordinary historical Evaluate continues to match the actual recording independently.

## Evidence and traceability

- Ordinary worksheet selection, baseline **54 PASS / 1 FAIL**, corrected owner **55/55**. The failed case selected Filtered Water and created Black Tea. The protecting test verifies the second output's SKU, fresh key, exact debit/credit receipts and unchanged unrelated work.
- Replay configuration, baseline **83 PASS / 1 FAIL** through Configure execution. The catalog 48 refusal in handoff 275 remains the predecessor result. Both new owner paths use a real guide exported from the source and imported into a separate Admin-generated Training destination.
- Stop after notification loss, worksheet candidate 02 **126 PASS / 1 FAIL**. The owner had already returned, but Stop waited for delivery of its notification. Reusable candidate 03 passed **120/120** after lifetime cleanup was separated from the notification.
- Final reusable Process **125/125**, controller `6d7b100cb08449aa8c62c55ae3bfa4e9`. The synthetic owner activity `62b6e79b-9071-4f86-b16b-94f7801381fb` produced Inventory events `CC52B08A-718F-4D74-965D-0BC5A603E6A8` and `FD317892-4129-411E-9B96-4D7B626B653F`. A separate two-step guide repeated the same control. Current proof for the first step released Next, which created a new binding for the second. Stop and Verify retained the first exact match and reported the second StepId missing.
- Final worksheet output **129/129**, controller `0ec5ccf19d514b80b6bf802ce43986b2`. The synthetic owner activity `ea2b8a82-1df6-4a17-955a-8d9c43517362` produced Inventory events `2BDB7B6C-7C32-4F97-88C2-A378C80D3B54` and `9300366D-25CA-48C1-AA6B-2F1EF158C139`. Each owner test independently read persisted inventory, the fresh output, exact input/output balance changes, both receipts and unchanged unrelated entities. The worksheet case debited 10 from the selected input, credited 1 to the new output and wrote its Production log.
- In both final cohorts an earlier, already-published ordinary completion in the same recording remained unclaimed. A different workbook completed normally without credit. Clearing the required actual quantity invalidated the selection, and the ordinary owner recorded REJECTED with zero inventory change. Check blocked that run. Changed session or warehouse, Stop, Close/reopen, an older publication, duplicate checks and two read-only Verify invocations preserved their expected history. Historical Evaluate still returned Awaiting on the older publication.
- Machine-local `production-loaded-packages.json` and `operator-completion-identities.json` sit under `reports/runtime/slice4be-production-complete-paths-reusable/52320fb1669447ebada0e924372e8859` and `reports/runtime/slice4be-production-complete-paths-worksheet/d40cd44213574cb0a22aec4496c72845`. These record the exact disposable loaded paths and their on-disk SHA-256 values, candidate paths/hashes, package identities and generated RunId, SequenceId, StepId, BindingId, owner activity and events. The table above pins pristine candidate bytes.
- On the final five, the ordinary worksheet test passed **55/55**. The shared runner and evidence changes also triggered **314/314** Seed regressions. Normal/interrupted two-step execution passed 18/18 and 22/22, nested Start/reopen 25/25 each, lost acknowledgment 22/22, the fixed Admin bridge/refusal group 21/21 and ordinary Seed owner boundaries 181/181.
- Receiving B0 passed **122/122** on the final candidate, controller `4221c1e10fdf4605823fc21ce240a84f`. The first run passed 116 and failed five layout checks because the old probe required both Production-only buttons to be visible in Receiving. The corrected probe checks that both are hidden, requires every original Receiving control to be visible, and checks bounds and overlap even for the two hidden controls. Scripted minimum/default/enlarged and native maximize/restore all passed alongside Start, Next, Stop, authorization and busy-close behavior.
- The first broad ordinary Production activity regression passed 432 and failed ten exact-context assertions. All ten compared newly authored records to catalog 28, which introduced this control. The fixture now compares new records to the loaded Core catalog and still tests the original catalog 27/28 boundary. After the failed run, Excel persisted beyond Quit. Capture found no process-owned Excel windows. A targeted cleanup attempt returned access denied, then the process exited and the controller restored settings and verified package hashes. This run is excluded from qualification. An ordinary activity rerun with Excel visible passed the ten context checks but its screenshot probe refused an obscured Production form (413 passes before the capture failure). The fixture retained its original host configuration.
- The final ordinary Production regression passed **442/442**, controller `7977a6ee18e5421daa7fa7592d6ecf5a`. Both owner branches retained the historical catalog boundary, exact live recording context, paired activities and source identities, guard and permission behavior, and tracking notices under unavailable storage. With Excel still enumerated five minutes after Quit, the owner requested manual closure. The process-owned window capture showed no Excel window or modal. Windows refused termination with access denied. A successful read-only GetExitCodeProcess query returned 0 during teardown.
- The earlier hidden-host reusable run reported **115 PASS / 2 FAIL**. Verify saved two Concluded evaluations while the test expected the internal state word in the rendered caption. The product displays “Conclusion observed.” The isolated Excel process remained headless after Quit, required assisted cleanup, and eventually disappeared. Its controller restored settings and verified package preservation. Only separate naturally closed runs count toward qualification.
- Static comparison against the archived source: 393 → 396 components, 6799 → 6834 procedures, 147102 → 147701 lines, 194 → 193 duplicate groups. Literal Application.Run targets remain 10 and unresolved calls remain 45. None of the 27 existing oversized modules grew. New modules max 148 lines and new procedures max 51. Earlier text-based 4x and 4aj checks reproduce identical pre-change findings (8/10 and 6/7 respectively).

The release evidence comprises **1,187 passing checks in 12 final cohorts** on the same five pristine candidate packages. All 60 controller package path/hash comparisons matched. Each controller verified restored settings and closed Excel. The final ordinary process exited with Windows status 0 and the process entry cleared. No dialog was dismissed in these qualifying runs. The preservation audit matched all 15 installed/prior package hashes, all five original startup registrations, six archived Seed and eleven archived Production/Core source pins, and the staged candidate source. At 23:31:54 UTC there were zero Excel processes or test controllers and zero Excel Application events 1000/1001/1002 since 19:50 UTC. The candidate remains isolated. No packages were promoted.

## Do Not Repeat

Keep the ordinary owner's event attribution and the loaded-publication check. A matching SKU, workbook name, earlier ActivityId, or successful form return cannot identify this step. Preserve the six archived source overrides when building this catalog 48 candidate. Use the visible test host for the operator-binding matrix. The broad ordinary activity cohort captures a real Production form against its original hidden Excel host. Test path controls using disposable saved probe copies.

## Assumptions to re-verify

Confirm current branches, exact candidate identities, registration, settings and absence of Excel/controllers before another build. Preserve the mixed-package VBE crash as unexplained.

## Open questions and blockers

Visible Process and worksheet walkthroughs remain open. Ask the operator to identify the selection shown after Bind, find Refresh Events after an older publication and recognize when Next becomes available. Observe which window receives input after each click. The Seed acceptance reservation for workbook windows flashing in front remains open. Production Check In has no registered execution input. Formal A, B2, B3 and combined Release 1 acceptance remain open.

## Immediate next action

Use the ordinary Production Check In handler to classify `PRODUCTION_RUN_CHECK_IN` against a fresh Training destination before extending replay coverage.

## Critical references

Architecture v4.11 D18-REPLAY-01 and the ExecutionRun v2 wire definition, Plan 022 remaining-scope matrix, Controls, handoffs 270–275, `modExecutionRun`, `modOperatorStepEvidence`, `modProductionOperatorBinding`, `modProductionCompleteActions` and `Test-Slice4beProductionRunLocal.ps1 -CompleteBaselineOnly -CompletePathsOnly -CompleteOperatorBinding`.
