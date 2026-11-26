# 273 B1 nested entry while Seed is returning

## Goal and release outcome

Qualify a second Start and a second Run How-To while a real built-in Seed call is returning. Architecture v4.11 D18-REPLAY-01 items 4 and 6 require a guarded dispatch and independently attributable evidence. Both existing product guards passed the new tests.

## Current verified state

Last verified November 26, 2026. Test commit `3a343af0f08ab185efaffc7aeef3526d309b1602` on `work/d18-picker-save-repro` follows handoff 272's `44e3bff2`. The candidate's product source is `b2865aefe89c0fc9648834221228cc7f23f3d9d9`. Plan 022 and Controls are at `24e121362a8542d5029e3f6db5b0b8a8d6d27a86` on `work/perplexity-handoff218`, published as `work/d18-picker-save-repro-docs`. The installed review set remains separate. Handoffs 270 and 272 identify the same catalog 48 candidate in `deploy/validation-b1-seed-bridge-01`. Each new case checked the five loaded disposable paths and the identities in the saved run.

| R1-5 package | Build identity | Candidate SHA-256 |
|---|---|---|
| Admin | `63b7b7846f894fd7acef45ee00629ca4` | `D9A7DD57E9DDF76556FE9860ADD67953350404844FE236DC2B64E28D21FB2A05` |
| Core | `6613640618d445ba970c7a7057dbc834` | `50126214113788BF1850D949A5A3EFD5DC63A9D705F8130398AA47D0B6E2ACFA` |
| Designs.Domain | `b4e0e83b5d5c4aacb322a15e63154860` | `140D96BA938D25FAF0EE8067588E7845015C33F702635C298C3FDC5683D8E955` |
| Inventory.Domain | `75d677c146c04fb7b8435a41e83c163f` | `8C33665F30A4C0FF46C7BD987C27084FD9071D7E9672EE5FBF8C556D9EA7B71A` |
| Operations | `439cdc79bad5422297e3d3a7972ebbde` | `755FC3C2D8908F964A0455DAAC5F28CBBB73F84623D36F18FA23F3D104FF66ED` |

The code worktree has an existing untracked `tests/tooling/reports/`. The primary code checkout's edited instructions and the docs tree's prior instruction, guidance and handoff 067 changes are outside these commits. The historical code PR 17 and docs PR 8 are closed against `work/d14-packaged-chain-coverage` and `work/excel-shutdown-timing`, respectively. Publication uses the feature branches.

## Decisions and constraints

The disposable Admin bridge receives the ordinary owner's successful Boolean in `ApplyBuiltIn`. A synchronous, one-shot test callback at the following line takes the saved-state copy and attempts the nested UI action. The Admin call and Operations dispatch are both on the stack. The tap counts every bridge call, accepted claim and owner application independently of the callback. The two gestures use separate generated Training destinations. No added timer or event pump, synthetic business effect or product entry point is involved.

## Evidence and traceability

| Action during the first Seed | Observed entry | Nested attempts | Bridge calls / accepted claims / applications | Result |
|---|---|---:|---|---|
| Start Run | The button returned DISABLED. A separate call to the actual Click handler reached the executing guard once and returned. | 1 | 1 / 1 / 1 | 25/25, controller `47c258375f73483d80864875f47f96ca` |
| Published guides → Run How-To | The enabled button delivered its Click. After selection validation the handler reached the existing-runner guard once and returned. | 1 | 1 / 1 / 1 | 25/25, controller `8ff765bb3eaa4196a85d40ce66a9ffe2` |

Both cases kept the same visible, busy runner in its library, the same setup token, context, destination, guide/profile, run revision, claimed pending step and active recording. The saved boundary had one Running run with zero finished steps and two observations, REQUESTED and APPLIED. On normal return, Operations accepted one acknowledgment. The sole saved step completed with the ordinary owner's APPLIED activity and event. The sealed recording retained the same two observation IDs.

An independent read of saved Inventory found one additional applied receipt, 21 fresh destination keys with their expected catalog, quantity, location, condition and log values, and all three bootstrap groups intact. After ordinary Viewer refresh, Verify concluded from exactly that event and key set. Run and recording hashes matched across Verify. Close and reopen returned Ready without another dispatch, inventory change or training artifact change. Source-guide and source-inventory pins matched.

Reproduce with `tests/tooling/Test-Slice4beAdminSeedOutcome.ps1 -DeployRoot deploy/validation-b1-seed-bridge-01 -Phase GREEN`, separately adding `-SeedReplayNested Start` and `-SeedReplayNested Reopen`. The first instrumentation attempt failed to locate the assignment using case-sensitive VBA source matching. A unique case-insensitive match corrected the fixture before behavioral qualification. There was no qualified product RED.

Exceptional cleanup passed **22/22** with `-SeedReplayLostAck`, controller `4bdb0ee3e93e49908b874ded8dbaa103`. After a real owner application, Operations accepted zero successful responses and saved Failed / `OWNER_NOT_COMPLETED`. Its Failed step retained `APPLIED` and the event. Verify was enabled at return, concluded from the actual receipt and kept the failed history. Reopen enabled Start and produced no automatic retry. Normal and exceptional dispatch both exit through the runner's busy-release and control-update code.

The retained bridge suite (`-SeedReplayRun`) passed **21/21**, controller `52457c2af434495caa2613aa3ddcad97`. It covers the fresh, stopped, ordinary-repeat and replay-repeat outcomes plus altered token/step/kit, duplicate claim, stale authority, changed target, missing Admin and source preservation. Receiving B0's prior **121/121** applies to this same product candidate. Only test scripts changed, so the shared runner, launcher and busy-close paths retained their existing qualification.

The regenerated static reports on the catalog 48 qualification source matched the baseline at 393 components, 6,799 procedures, 147,102 lines, 194 duplicate-body candidates, 10 literal `Application.Run` sites and 45 unresolved calls.

The ordinary Admin Seed workflow, edge cases and recording boundaries passed **181/181** with `-Recording -RecordingBoundaries -WorkflowEdges`, controller `a5f1cf4977634e9389c483d1e4271193`. This retains the ordinary business owner and recorder behavior used by the bridge. The five qualifying runs passed **274/274** checks in total. Each compiled all five disposable VBA projects.

All six controllers, including the initial fixture failure, closed Excel naturally, restored settings and preserved their package and tooling hashes. No modal required intervention. The final audit matched six product-source pins, all 15 XLAM hashes across the candidate, preceding and installed sets, and the five normal startup registrations. It found zero Excel processes, test controllers or Excel Application error/hang/reporting events (1000/1001/1002) from 07:51:57 through 08:13:21 UTC on November 26. Local evidence is in `reports/runtime/b1-nested-entry-20261126`. Excel was closed at completion.

## Do Not Repeat

A direct second call to Admin only tests the claim backstop. A reopen after completion only tests subsequent access. Keep the original call on the stack for these two UI tests. Handoff 269's earlier mixed-package VBE crash remains unexplained.

## Assumptions to re-verify

Before another Excel qualification, verify package identities, disposable loaded paths, desktop/Excel ownership, settings and Training binding. The nested tests are packaged handler evidence. Justin's visible acceptance is for ordinary fresh and repeat Seed, with the workbook-flashing reservation. This qualification collected no foreground-window trace and establishes no click attribution or focus verdict.

## Open questions and blockers

The two qualified gestures cover single-step Seed. Broader adapter and partial-outcome coverage, B2 executable transfer/local results, retained 4be-A and B3 combined technical and visible Operations/Admin gates remain open. Trace workbook flashes separately. [Plan 022 remaining-scope matrix](../expert%20guidance%20docs/022%20Deployed%20Operations%20Launcher%20and%20NAS%20Runtime%20Stabilization%20Plan.md#remaining-4be-b-scope-against-the-milestone-exits)

## Immediate next action

On a fresh Training destination, qualify an already-bound multi-step guide with a later eligible action, lose the first applied Seed acknowledgment, and prove zero dispatch of the later action while preserving the first event and failed run history.

## Critical references

`frmActionPathRun.mStart_Click/DispatchSteps`, `frmActionPathLibrary.mRunHowTo_Click`, `modExecutionRun.OpenSetup/PendingSeed/FinishStep/Verify`, `modAdminSeedReplay.ApplyBuiltIn`, `modAdminSeedObservation.Seed`, `Slice4beSeedNestedEntry.ps1`, `Slice4beSeedLostAck.ps1`, Plan 022 B0–B3, Controls v1 and handoffs 269–272.
