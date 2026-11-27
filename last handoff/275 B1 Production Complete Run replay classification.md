# 275 B1 Production Complete Run replay classification

## Goal and release outcome

Classify `PRODUCTION_RUN_COMPLETE` under D18-REPLAY-01 items 2–6. Branch B applies on catalog 48. The ordinary control has a recorded owner and supports historical verification. As a required execution step it is an **unsupported deliberate action**. D18 item 5 keeps this coverage obligation open. The proposed milestone treatment is **operator-required**, pending a decision and a qualified way to associate the operator's action with the pending destination step.

## Current verified state

Last verified November 27, 2026. The starting test and documentation commits were `c10eb7a986a185f338038b4ac25603fe0c79ed23` and `ffe5c9ff88d3a8aef1c30e814365aba6a21eac12`. The source for the isolated catalog 48 candidate is `b2865aefe89c0fc9648834221228cc7f23f3d9d9`. Eleven reviewed Production and Core source files matched the qualification stage. Test copies came from `deploy/validation-b1-seed-bridge-01`.

Test commit `d8e5ad0ee531c9dffdc9c6529585218c7453dbeb` is on `work/d18-picker-save-repro`. Plan 022 and Controls commit `88dd851f5d3cc260190210120b20709b3b57d2b2` is on `work/perplexity-handoff218`, published as `work/d18-picker-save-repro-docs`. Excel closed naturally. The qualification leaves the configured five startup add-ins in place. Unrelated existing changes include the main checkout's instructions, the test checkout's untracked tooling reports, and the docs checkout's deleted instructions, older handoff and guidance changes and untracked files.

| R1-5 package | Build identity | Original candidate SHA-256 |
|---|---|---|
| Admin | `63b7b7846f894fd7acef45ee00629ca4` | `D9A7DD57E9DDF76556FE9860ADD67953350404844FE236DC2B64E28D21FB2A05` |
| Core | `6613640618d445ba970c7a7057dbc834` | `50126214113788BF1850D949A5A3EFD5DC63A9D705F8130398AA47D0B6E2ACFA` |
| Designs.Domain | `b4e0e83b5d5c4aacb322a15e63154860` | `140D96BA938D25FAF0EE8067588E7845015C33F702635C298C3FDC5683D8E955` |
| Inventory.Domain | `75d677c146c04fb7b8435a41e83c163f` | `8C33665F30A4C0FF46C7BD987C27084FD9071D7E9672EE5FBF8C556D9EA7B71A` |
| Operations | `439cdc79bad5422297e3d3a7972ebbde` | `755FC3C2D8908F964A0455DAAC5F28CBBB73F84623D36F18FA23F3D104FF66ED` |

## Decisions and constraints

| Question | Ordinary contract and replay implication |
|---|---|
| What was recorded? | Production Run - List → Complete Run calls `frmProduction.mBtnManagerApplyOutput_Click` → `modProductionCompleteActions.Execute` → `frmProduction.CompleteProductionRun`. Catalog owner `PRODUCTION_RUN_COMPLETION` covers the selected reusable Process or worksheet output. The wrapper handles busy/loading and observation. The two business owners below perform the writes. A recording identifies the control, activity, actor, warehouse, sequence, outcome and exact event references. It omits business identities and entered quantities. |
| What travels? | Guide intent includes the ordered control, instructions and expected outcome/conclusion. Source activity and event IDs describe provenance. Destination workbook instance, actor/session, recipe revision, batch, node, allocations, output selection, actuals, session/event IDs and System_Keys must be local. |
| Who may act? | `cProductionWorksheetAction.Begin/CanContinue` uses `modProductionDesignerActions.ContextIsCurrent` to check the captured context and the same open Workbook object, then checks PROD_POST or ADMIN_MAINT. It rechecks during yielding work. The worksheet owner additionally requires PROD_POST. A guide gives no Production rights. Any replay would also require a generated, bound Training authority and compatible profile, policy and package identities. |
| What makes a reusable Process eligible? | Load and validate the released recipe version/graph with its Process definitions, select the Process in the local batch, prepare the location and required upstream outputs, allocate exact available inputs and Check In that node. `CompleteReusableProcess` requires `mCheckedIn` for that specific `ProcessNodeId`, rechecks readiness and live allocations, and validates actual quantities for its outputs. The form stages the selected active output. The first completion creates `mRunId` inside the owner. It is unavailable as a pre-completion selector for that first node. |
| What makes worksheet completion eligible? | In the captured workbook, checked-in `Prod_invSys_Check` rows carry exact input keys and used quantities. The owner requires valid allocation locations, a resolvable selected `ProductionOutput` row and a positive numeric Actual Output. Where preparation is needed, an ordinary recipe OUTPUT supplies staging. Existing output keys must resolve uniquely. Blank-key draft selection can use local Process/output fields or row position. `CreateProductionSessionFromWorkbook` creates the typed session and output/event identities during completion and stores session metadata on the workbook. No portable pre-existing run ID is specified. |
| How is completion confirmed? | The deliberate click invokes the selected owner; this path has no generic Yes/No replay confirmation or distinct CANCELLED completion outcome. Selection or input refusal can return REJECTED. Initial authorization refusal is DENIED. A submission/refresh failure can raise an ordinary worksheet notice. Preserve these owner boundaries. |
| What persisted? | Reusable `CompleteReusableProcess` queues applicable PROD_CONSUME and PROD_COMPLETE events, invokes the processor, verifies node output balances and records node completion/history. It marks the batch complete only after all nodes finish. Worksheet `mProduction.CompleteProductionRunAfterCheckInForOutput` calls `modProductionCompletionService.ExecuteProductionSession`, queues the consume and complete events, processes and refreshes, writes ProductionLog and clears managed staging after success. Inventory owns applied entities and receipts. |
| What proves this invocation? | Pair REQUESTED and the owner's terminal record for the same current activity and sequence. CONFIRMED proves command acknowledgment and completed local refresh. SourceEventsApplied additionally requires **every** referenced event to be applied in the owning publication. Inspect saved receipts, exact input/output keys and line deltas; for worksheet completion also inspect the local log. The samples use one consumed input and one fresh output. Earlier receipts, aggregate stock and a form returning cannot identify a new completion. |
| Repeat or uncertain outcome? | A completed reusable node cannot be checked in again in that batch; success clears its check-in binding. Successful worksheet completion clears managed check rows and output entry. A repeat without new preparation has no fresh completion evidence. The writer reports Submitted for acknowledged submissions and Unknown for an uncertain attempted write. PENDING can retain acknowledged work awaiting processing/refresh; FAILED can retain partial or uncertain references. Context loss can leave only the original attempt. Inspect all exact references. Subsequent dispatch, automatic retry and rollback have no authorization. |
| Can a profile bind it now? | `modExecutionProfileModel.NewSteps` requires every control to pass `modExecutionInputs.Supported`. The seven registered controls have no Production workbook, batch/node or staged-output selector. `OpenDraft` refuses the whole required profile; `modExecutionBinding.Read` requires a profile for the exact guide version. Neither source IDs nor the currently active workbook, first matching name or latest run provide an approved mapping. |

## Evidence and traceability

The focused test exercises Configure execution and Run How-To on a locally authored guide after genuine ordinary completion evidence already exists. Disposable counters cover the form click and entry to its completion routine. The guide author configures the guide and a signed-in operator with PROD_POST attempts to run it. The generated Training authority is local to the test. The existing isolated target fixture supplies a synthetic UNC hub, so the result covers packaged behavior only. The guide fixture separately checks publication freshness, independent source and observed recordings, exact Inventory events, guide integrity and historical evaluation. Its worksheet variant prepares an isolated operator surface and calls the actual Check In and Complete Run controls. Setup and ordinary recording supply the historical examples.

Run each case separately with `tests/tooling/Test-Slice4beProductionRunLocal.ps1 -DeployRoot deploy/validation-b1-seed-bridge-01 -Phase GREEN -CompleteBaselineOnly -CompletePathsOnly -CompleteReplayRefusal -SavedProbeCopiesForTest -CompletePathMode Reusable`, substituting `Worksheet` for the second case. The workbook and add-in paths are disposable copies. Local evidence is under `reports/runtime/b1-production-classification-20261127` and `reports/runtime/production-run-local-controller`.

Three initial test runs exposed fixture assumptions. The first stopped at the build-property probe (43 passed, 1 failed), the second compared a current observation with hardcoded catalog 28 (53 passed, 1 failed), and the third attempted Run How-To in the guide-author session (100 passed, 1 failed). In the third run Configure gave the correct refusal; its Run result did not establish the authorized operator's missing-profile response. The final fixture reads the actual loaded build properties and catalog, and separately verifies the Production operator's ordinary capability. These were test corrections. No behavioral RED was established on the candidate.

| Final packaged case | Checks passed | Evidence controller |
|---|---:|---|
| Reusable Process | 101/101 | `43a3eb7a164f4f77939d73384341ef90` |
| Worksheet output | 107/107 | `955e4f3a7bc94d2bb13f8b91a3e8f8e7` |

Both cases compiled all five disposable projects. Configure refused the required unsupported step and opened no profile editor. Run How-To reported the missing exact guide profile and opened no runner. Each pair of counters stayed at two ordinary clicks and two owner entries. Journal, profile/run file inventory, publication and business authority hashes were preserved. Historical evaluation of the separate second recording credited only its own two exact events. Its older publication produced Awaiting.

The regenerated static scan matched the candidate baseline of 393 components, 6,799 procedures, 147,102 lines, 194 duplicate bodies, 10 literal `Application.Run` sites and 45 unresolved calls. Across all five runs, Excel closed naturally, settings were restored and the packages were preserved. No modal was dismissed. The final audit matched all 15 package hashes, six baseline and eleven reviewed source pins, and the five ordinary startup registrations. It found zero Excel processes, controllers or Excel Application events 1000/1001/1002 from 19:14:34 through 19:42:06 UTC on November 27. No XLAM rebuild, reload or promotion is required for this test and documentation change.

Catalog 48 cannot create a destination execution binding for this control. The applicable stale-evidence check uses an older publication against newly generated event IDs. A stale bound-run execution case and a replay retry require a supported binding first. The historical ordinary-owner regression includes a second click after success and verifies that inventory does not change again. The prior eight Seed, Receiving and interruption gates passed 435/435 on these same five candidate binaries (handoff 274). This increment touches test tooling and documentation only. The focused cases exercise both ordinary Production owner branches, recording, authorization and historical evaluation. No shared dispatch, progression, binding or owner code changed to trigger wider regressions.

## Do Not Repeat

Do not treat the first local Process name match, an unkeyed staging-row position, a source event or a session ID created during completion as a destination identity. A successful local CONFIRMED observation alone cannot establish Domain application. Handoff 269's mixed-package VBE crash remains unexplained.

## Assumptions to re-verify

Recheck original candidate hashes, disposable loaded paths, desktop and Excel ownership, settings and Training context before another run. Justin's visible acceptance applies to ordinary fresh/repeat Seed with a workbook-flashing reservation. These Production checks do not qualify foreground focus.

## Open questions and blockers

Should Complete Run use operator-required execution with an explicitly captured local Production selection, or should a future automatic adapter be designed? For operator-required execution, specify how the runner freezes and revalidates the destination branch, workbook, batch/node or worksheet target, actor and step and associates the ensuing fresh owner evidence. An automatic adapter also needs reviewed typed inputs for actual outputs and a safe identity before the first completion. Both branches need a stale-selection refusal and an already-completed negative control. Production Check In (`PRODUCTION_RUN_CHECK_IN`) is a separate observed, unregistered prerequisite. B2, B3, formal 4be-A and combined Release 1 acceptance remain open.

## Immediate next action

Decide the intended operator-required binding for the prepared destination selection before authorizing an execution adapter or a new profile input.

## Critical references

In the docs repository, see Architecture v4.11 D18-REPLAY-01 items 2–6 and “Complete Run observation refinement”, the Plan 022 remaining-scope matrix, Controls v1 and handoff 274.

In the code repository, the entry and both owners are in `src/Production/Forms/frmProduction.frm` and `src/Production/Modules/{modProductionCompleteActions,modProductionReusableRun,mProduction,modProductionCompletionService}.bas`. The captured workbook and context guards are in `src/Production/ClassModules/cProductionWorksheetAction.cls` and `src/Production/Modules/modProductionDesignerActions.bas`. The published guide buttons are in `src/Operations/Forms/frmActionPathLibrary.frm`. Under `src/Core/Modules/`, see `modProductionCompleteCodes`, `modExecutionProfileModel.NewSteps`, `modExecutionInputs.Supported` and `modExecutionBinding.Read`. The packaged proof is in `tests/tooling/Slice4beProductionCompletePaths.ps1`. Older owner interruption and repeat evidence is in `tests/integration/plan022_slice4be_production_complete_activity_results.md`.
