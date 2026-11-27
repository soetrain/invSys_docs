# 274 B1 multi-step interruption after Seed acknowledgment loss

## Goal and release outcome

Qualify D18-REPLAY-01 items 4 and 6 for Run all when Seed applies, Operations loses its acknowledgment and a second required action is eligible. The existing runner stopped before Open Receiving. Verify retained the Seed effect and failed the whole guide for the missing step.

## Current verified state

Last verified November 27, 2026. Test commit `c10eb7a986a185f338038b4ac25603fe0c79ed23` on `work/d18-picker-save-repro` follows `3a343af0`. Plan 022 and Controls are at `cb3ec80526246efb4a1bdef99d22b7b430cca556` on `work/perplexity-handoff218`, published as `work/d18-picker-save-repro-docs`. The tests use disposable copies of the handoff 270 catalog 48 candidate, `deploy/validation-b1-seed-bridge-01`, product source `b2865aefe89c0fc9648834221228cc7f23f3d9d9`. Both focused cases compiled all five projects, checked the loaded paths and matched each saved run's package identities.

| R1-5 package | Build identity | Original candidate SHA-256 |
|---|---|---|
| Admin | `63b7b7846f894fd7acef45ee00629ca4` | `D9A7DD57E9DDF76556FE9860ADD67953350404844FE236DC2B64E28D21FB2A05` |
| Core | `6613640618d445ba970c7a7057dbc834` | `50126214113788BF1850D949A5A3EFD5DC63A9D705F8130398AA47D0B6E2ACFA` |
| Designs.Domain | `b4e0e83b5d5c4aacb322a15e63154860` | `140D96BA938D25FAF0EE8067588E7845015C33F702635C298C3FDC5683D8E955` |
| Inventory.Domain | `75d677c146c04fb7b8435a41e83c163f` | `8C33665F30A4C0FF46C7BD987C27084FD9071D7E9672EE5FBF8C556D9EA7B71A` |
| Operations | `439cdc79bad5422297e3d3a7972ebbde` | `755FC3C2D8908F964A0455DAAC5F28CBBB73F84623D36F18FA23F3D104FF66ED` |

The installed, preceding and isolated candidate sets have separate identities. Existing local changes in the primary code checkout, untracked tooling reports and unrelated docs instructions, guidance and handoff 067 remain outside these commits. The historical PRs remain closed with their original bases. [Code PR 17](https://github.com/soetrain/invSys_fork/pull/17) · [Docs PR 8](https://github.com/soetrain/invSys_docs/pull/8)

## Decisions and constraints

The source recording uses the ordinary Admin Seed and Receiving Open owners. Each case exports its authored guide, imports into a separate generated Training destination, and saves a local two-step profile bound by guide hash and stable StepIds. The actor has `ADMIN_MAINT`, `ACTION_PATH_MAINT` and `RECEIVE_POST`. The runner requires explicit selection of an exact destination bootstrap System_Key for this mixed guide. Both required outcomes are ordered. The authored source-events conclusion selects the first step so the evaluator can retain the applied Seed event even if the later requirement is missing.

Only disposable packages contain the counters. The two Receiving counters bracket the typed runner dispatch and actual `ShowReceivingForm` entry and are armed after authoring. The normal control uses the original Seed return. The fault case reuses handoff 272's exact boundary. It copies the saved state after a successful real return, then raises before Operations accepts the Boolean. The runner enters `FinishStep(False)` and performs its ordinary cleanup. No manual Stop is delivered.

## Evidence and traceability

| Measured boundary | Normal control | Lost acknowledgment |
|---|---|---|
| Focused packaged gate | 18/18, controller `febcf6c5c5b64435bcebaed10c878c6a` | 22/22, controller `a78b6386467e41cb904f04227fea62d6` |
| Seed bridge / accepted claim / inventory application | 1 / 1 / 1 | 1 / 1 / 1 |
| Seed acknowledgments accepted / dropped / error returns | 1 / 0 / 0 | 0 / 1 / 1 |
| Receiving dispatch / owner calls | 1 / 1 | 0 / 0 |
| Saved receiving effect | Previously absent station-local workbook created, bound and opened with its ReceivedTally table. The closed journal retained REQUESTED and OPENED for the second activity. | Workbook remained absent. No Receiving step revision or recording pair appeared. |
| Run and Seed step | Completed; Seed Completed with APPLIED; Receiving Completed with OPENED. | Failed / `OWNER_NOT_COMPLETED`; Seed Failed with APPLIED and its original event. |
| Whole-guide Verify | Concluded, both required actions matched. | Failed / `REQUIRED_STEP_MISSING`; Seed matched, the expected Receiving step was missing. |

In each case the destination inventory had exactly one additional applied receipt, 21 new unique keys with matching kit values and event lines, and all three bootstrap groups retained their keys and values. The journal linked the same run, Seed activity and submitted destination event. Verify credited exactly that event and its 21 keys, distinct from the source recording. For the fault case, the Inventory hash matched the copy taken before Operations handled the lost acknowledgment.

The interrupted runner's Verify button produced two independent saved evaluations. After Close, reopening Run How-To required an entity choice and enabled Start after selection. The new setup had no attempt to Verify. The ordinary historical route selected the original closed recording, applied the guide and evaluated again. The third evaluation retained the same failed verdict, credited event and missing step. Run revisions, recording files, inventory and counters remained intact throughout. The source guides and source warehouses were preserved.

Reproduce using `tests/tooling/Test-Slice4beAdminSeedOutcome.ps1 -DeployRoot deploy/validation-b1-seed-bridge-01 -Phase GREEN` with `-SeedMultiStep Normal` and `-SeedMultiStep Lost` in separate processes. Both cases passed on the first packaged run. There was no behavioral RED and no change to product code, package registration or deployment.

Retained single-step acknowledgment loss passed **22/22** (`e9adee5e53544eb5bef4d0eb8a2223ac`). Nested Start and Run How-To each passed **25/25** (`f6a18625bf3e4df4a43c26e6807e244a`, `3682c129a85f4b77a65f5cb3eabdff02`). The original bridge and refusal suite passed **21/21** (`67dfd34e6dd24bdab88f8e75570b63b5`), retaining fresh/stopped/repeat attribution, altered and duplicate requests, stale target/authority and unavailable Admin. These protect the existing owner, progression and reentry contracts on the same candidate.

Static evidence regenerated from the qualified product source matched the existing baseline of 393 components, 6,799 procedures, 147,102 lines, 194 duplicate-body candidates, 10 literal `Application.Run` sites and 45 unresolved calls.

Ordinary Admin Seed, workflow edges and recording boundaries passed **181/181** (`4521a0b78cc445bebe54f179c815c505`). Receiving B0 passed **121/121** (`5642d4e1255d45a5bc5d6819c7773f7a`) with `Test-Slice4beReceivingReplay.ps1 -DeployRoot deploy/validation-b1-seed-bridge-01 -Phase GREEN -CaptureEvidence`. It independently exercised the selected later owner, its normal recording, target guards and busy-close handling.

The eight gates passed **435/435**. All closed Excel naturally, restored the settings snapshot and preserved package and tooling hashes. No modal was dismissed. The final audit matched all six product-source pins, all 15 candidate/preceding/installed XLAM hashes and the five ordinary startup registrations. It found zero Excel processes, test controllers and Excel Application error/hang/reporting events (1000/1001/1002) during 18:07:22–18:42:25 UTC on November 27. Excel was closed at completion. The local evidence directory is `reports/runtime/b1-seed-multistep-20261127`.

## Do Not Repeat

Matching stock alone cannot establish that Receiving never dispatched. The positive control creates its own workbook under a disposable operator root. The fault case starts with its own empty destination and uses both counters and saved evidence. The earlier mixed-package VBE crash in handoff 269 remains unexplained.

## Assumptions to re-verify

Before another live gate, check the five build IDs and hashes, loaded disposable paths, desktop/Excel ownership, ordinary registrations and Training binding. Justin accepted ordinary fresh and repeat Seed with a reservation about workbook windows flashing in front. These packaged tests collected no foreground-focus trace or new human acceptance.

## Open questions and blockers

The two-step case closes the specific gap in handoff 273. Production Complete Run has recorded ordinary owner evidence and lacks an entry in the current seven-control execution input registry. Its deliberate operation needs an explicit replay classification and prerequisite coverage. Other B1 adapters and boundaries, B2 executable transfer/local revalidation and Admin results, retained 4be-A and B3 combined technical and visible acceptance remain open. [Plan 022 remaining-scope matrix](../expert%20guidance%20docs/022%20Deployed%20Operations%20Launcher%20and%20NAS%20Runtime%20Stabilization%20Plan.md#remaining-4be-b-scope-against-the-milestone-exits)

## Immediate next action

Trace `PRODUCTION_RUN_COMPLETE` from its ordinary packaged handler through the control census and profile registry, and specify the focused test for its supported replay classification and prerequisites.

## Critical references

`frmActionPathRun.DispatchSteps/mVerify_Click`, `modExecutionRun.FinishStep/Verify`, `modReceivingExecution.Dispatch`, `modTS_Received.ShowReceivingForm`, `modEvaluationMatches`, `modPathEvaluation.Classify`, `Slice4beSeedMultiStep.ps1`, `Slice4beSeedLostAck.ps1`, `Slice4beSeedReplayProof.ps1`, Controls v1, handoffs 270–273.
