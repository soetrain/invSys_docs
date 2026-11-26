# 272 B1 Seed applied with acknowledgment lost

## Goal and release outcome

Qualify the single B1 interruption between an actual Admin inventory application and acceptance of its return by Operations. D18-REPLAY-01 items 4 and 6 govern dispatch and independent proof.

## Current verified state

Last verified November 25, 2026. Test tooling commit `44e3bff29e10a898016fde01fd0ea8682903a44f` is on `work/d18-picker-save-repro`. Product source remains at `b2865aefe89c0fc9648834221228cc7f23f3d9d9`. Starting docs were `4ad4954e60395b923e7d44bf03cdb42419717dd1` on `work/perplexity-handoff218`, published as `work/d18-picker-save-repro-docs`. The tests loaded isolated copies of the handoff 270 catalog 48 candidate from `deploy/validation-b1-seed-bridge-01`. The fixture checked loaded paths and run-bound build identities. SHA-256 below identifies the original candidate before disposable instrumentation.

| R1-5 package | Build identity | SHA-256 |
| --- | --- | --- |
| Admin | `63b7b7846f894fd7acef45ee00629ca4` | `D9A7DD57E9DDF76556FE9860ADD67953350404844FE236DC2B64E28D21FB2A05` |
| Core | `6613640618d445ba970c7a7057dbc834` | `50126214113788BF1850D949A5A3EFD5DC63A9D705F8130398AA47D0B6E2ACFA` |
| Designs.Domain | `b4e0e83b5d5c4aacb322a15e63154860` | `140D96BA938D25FAF0EE8067588E7845015C33F702635C298C3FDC5683D8E955` |
| Inventory.Domain | `75d677c146c04fb7b8435a41e83c163f` | `8C33665F30A4C0FF46C7BD987C27084FD9071D7E9672EE5FBF8C556D9EA7B71A` |
| Operations | `439cdc79bad5422297e3d3a7972ebbde` | `755FC3C2D8908F964A0455DAAC5F28CBBB73F84623D36F18FA23F3D104FF66ED` |

Plan 022 and Controls record the qualified behavior at `5ae9ece12fdce6dfd35c492a57c88b43e22c3ea2`.

## Decisions and constraints

In the disposable Operations form, a private function receives the real fixed Admin call's Boolean as an argument. After a successful return it copies the saved Inventory, journal and run revisions, then raises a controlled error before `delivered` can be assigned. The ordinary error branch calls `FinishStep(False)`. The normal build exposes no fault control. Admin instrumentation counts entry, passage through `PendingSeed(..., True)` and the ordinary Seed observer's APPLIED completion. There is no substituted owner. The processor's receipt and event remain the inventory authority.

## Evidence and traceability

- Reproduce with `tests/tooling/Test-Slice4beAdminSeedOutcome.ps1 -DeployRoot deploy/validation-b1-seed-bridge-01 -Phase GREEN -SeedReplayLostAck`. Final focused result **21/21**, controller `cd69be8a8389402ba30c834cbc85cedd`. All five disposable projects compiled.
- At the fault boundary the bound run was Running with zero steps. The saved recording already had REQUESTED and APPLIED for one activity, plus the destination's submitted Inventory event reference. The inventory copy already contained one additional APPLIED receipt and 21 additional log lines. Its hash matched Inventory after dispatch stopped. The three bootstrap groups retained their exact keys and row values. The other 21 matched their kit values, catalog values and log lines by fresh destination System_Key.
- Operations counters showed one bridge invocation, one dropped positive response, zero accepted successes and one error return. Admin counters showed one entry, one accepted claim and one ordinary replay application. Expected and observed terminal state were Failed/`OWNER_NOT_COMPLETED`; the sole step was Failed with owner reason `APPLIED`. The sealed closed journal retained both observation identities and linked to the pre-loss journal hash.
- After Viewer refresh, the actual Verify button returned Concluded. The evaluation bound the exact guide and closed recording and credited one Applied event with all 21 exact keys. No source-guide event was borrowed. Verify preserved the run revisions and recording bytes. Close and reopen from Published guides returned to the same Ready profile. Invocation and owner counters, the entire destination training artifact set and inventory hash were unchanged. The source guide and warehouse pins also matched.
- The first comparison ran **19 pass / 1 fail**. Both observations share the same action ordinal; sorting that field put APPLIED first in the diagnostic. Sorting the versioned journal records gave REQUESTED followed by APPLIED, with both record IDs matching the closed pair. The test now checks the pre-loss pair and its immutable close separately. This was a comparison defect in the test.
- Static scan of the same qualification source yielded 393 components, 6,799 procedures, 147,102 lines, 194 duplicate-body candidates, 10 literal Application.Run sites and 45 unresolved calls. The approved fixed bridge is unchanged.
- The retained bridge suite passed **21/21** on the original build identities, including independent fresh proof, stopped-run zero credit, ordinary and replayed no-op, altered and duplicate bridge requests, stale authority, changed destination, unavailable Admin and source preservation. Incompatible Admin path/build checks were reviewed in the unchanged binding code. Receiving's prior 121/121 belongs to this same candidate. No production runner, prerequisite, binding or add-in changed in this increment, so Receiving B0 and the four other unrelated gates were not rerun.
- Ordinary Admin Seed, workflow edges and recording boundaries passed **181/181**. Retained controllers were `73dff57c562247dca94ed2070fb7f642` (bridge) and `fa46e6e7501d46a4a4de964101cfa442` (ordinary Seed). All four runs, including the initial comparison failure, closed Excel naturally, restored settings and preserved their complete tooling and package pins. The final audit matched six original product-source pins, all 15 candidate/preceding/installed XLAM hashes and the five ordinary startup registrations. It found zero remaining Excel or test-controller processes and zero Excel Application error/hang/reporting events (1000/1001/1002) between 03:31:15 and 03:48:22 UTC on November 26. No modal was dismissed. The local qualification directory is `reports/runtime/b1-lost-ack-20261126`.

## Do Not Repeat

Order the two observations by journal version. Taking a snapshot after Verify cannot establish which information existed when the acknowledgment was lost. Handoff 269's mixed-package VBE crash remains unexplained.

## Assumptions to re-verify

Before another packaged run, check Excel and controller ownership, the desktop session, the original five package hashes, disposable copy paths and ordinary startup registrations.

## Open questions and blockers

| Scope | Status |
| --- | --- |
| Fresh and repeat built-in Seed | Technically qualified; Justin accepted the visible workflow with a reservation about frequent workbook flashes. |
| Lost acknowledgment after Seed applies | Technically qualified through Start, Verify and reopen. |
| Nested entry | Proposed next increment. Two UI gestures still need tests while the original owner is returning. |
| Workbook flashes | Separate open UI finding. Trace the exact click, foreground workbook/window and input focus throughout the action. Distinguish product workbook opens from setup or diagnostic reads and check for an obscured prompt. |
| Broader B1, 4be-B, formal 4be-A and Release 1 | Open. No package promotion. |

## Immediate next action

For nested entry, use two independent, one-shot test-owned gestures in disposable builds while the first real Seed call is still on the stack, immediately after the real owner applies and before its return reaches Operations. In one case attempt Start again through the existing button; observe the disabled control or busy guard. In the other attempt Run How-To through its Published guides action; observe the existing runner reference and Core's pending-step guard. Count the second UI attempt even if the disabled button delivers no handler. Each case should retain one form/run identity, one bridge invocation, one accepted claim, one owner event and one receipt. The outer Start should finish normally and Verify should credit only that event. Keep nested entry separate from a deliberately dropped acknowledgment.

## Critical references

Plan 022 B1; Controls v1; handoffs 269–271; `frmActionPathRun.DispatchSteps`; `frmActionPathLibrary.mRunHowTo_Click`; `modAdminSeedReplay.ApplyBuiltIn`; `modAdminSeedObservation.Seed`; `modAdminInventorySeed.QueueDemoCreateAndProcess`; `modExecutionRun.PendingSeed/FinishStep/Verify`; `Slice4beSeedLostAck.ps1`; `Slice4beSeedReplayProof.ps1`.
