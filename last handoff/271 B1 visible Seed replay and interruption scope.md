# 271 B1 visible Seed replay and interruption scope

## Goal and release outcome

Review the built-in B1 Seed How-To with an operator on the qualified candidate and scope a test of a lost dispatch acknowledgment under D18-REPLAY-01.

## Current verified state

Last verified November 25, 2026. Code `b2865aefe89c0fc9648834221228cc7f23f3d9d9`, branch `work/d18-picker-save-repro`. Starting docs `2557cdd993c04467407824bafefd89c823132c96`, branch `work/perplexity-handoff218`, published as `work/d18-picker-save-repro-docs`. This attended review used `deploy/validation-b1-seed-bridge-01`. All five open add-in paths and hashes matched handoff 270. Both saved runs bound those five R1-5 builds. Handoff 270 retains the prior seven technical gates (21, 121, 181, 110, 259, 59 and 33 checks respectively).

| Package | Loaded build identity | SHA-256 |
| --- | --- | --- |
| Admin | `63b7b7846f894fd7acef45ee00629ca4` | `D9A7DD57E9DDF76556FE9860ADD67953350404844FE236DC2B64E28D21FB2A05` |
| Core | `6613640618d445ba970c7a7057dbc834` | `50126214113788BF1850D949A5A3EFD5DC63A9D705F8130398AA47D0B6E2ACFA` |
| Designs.Domain | `b4e0e83b5d5c4aacb322a15e63154860` | `140D96BA938D25FAF0EE8067588E7845015C33F702635C298C3FDC5683D8E955` |
| Inventory.Domain | `75d677c146c04fb7b8435a41e83c163f` | `8C33665F30A4C0FF46C7BD987C27084FD9071D7E9672EE5FBF8C556D9EA7B71A` |
| Operations | `439cdc79bad5422297e3d3a7972ebbde` | `755FC3C2D8908F964A0455DAAC5F28CBBB73F84623D36F18FA23F3D104FF66ED` |

## Decisions and constraints

Justin's verdict after the fresh run, refresh and repeat was **“i think its fine for now, accepted”**. It applies to the visible built-in Seed workflow on this candidate. He supplied the repeat screen wording before giving the verdict. The display finding is recorded in Controls. Built-in Seed is technically qualified and visibly accepted. Broader B1 stays open and the formal 4be-A and Release 1 decisions carry forward.

The operator signed in through the ordinary Admin prompt. The signed-in Create Warehouse command created a separate Training source and destination on the designated review share. Preparation used scoped Users & Roles, capture policy and snapshot commands, an ordinary source Seed recording, guide authoring/export and destination import/configuration. A transient review workbook drove setup through the packaged services. The operator navigated the destination Viewer, selected the imported How-To and ran each attempt through its form. The installed reader-reopen packages retained their startup registration.

## Evidence and traceability

- The destination had three bootstrap groups before the first Start. The exact imported guide and local profile bound the built-in kit. A single Start without a Receiving selection reached ordinary Seed. The closed recording paired REQUESTED and APPLIED with the current activity. An independent read of saved Inventory matched the receipt and 21 log lines to 21 new destination keys and preserved the three bootstrap groups.
- Direct Verify initially displayed Awaiting and referenced the current submitted event. At the reviewer's direction, the operator brought Viewer forward, refreshed Events and verified again. The result was Concluded with that exact applied event and 21 lines. This extra interaction is part of the visible result. Publication was already loaded when the repeat was verified.
- A new explicit Start returned UNCHANGED. The second closed recording had its own activity pair. Stock, receipt and log counts showed no new effect. Verify returned Failed with REQUIRED_OUTCOME_MISMATCH and credited zero terminal sources. The operator reported seeing “Seed Demo Inventory: Outcome mismatch. Expected APPLIED.” The reported text omitted the observed UNCHANGED outcome and the reason no fresh Inventory event was expected. The saved run revisions spanned 113.35 seconds for fresh Seed and 26.47 seconds for repeat. These are NAS run intervals, excluding navigation and Verify.
- Both attempts bound the same destination, guide and profile. The original source inventory, sealed source guide and export hashes matched after the repeat. The private review results contain boolean checks and counts. No inventory keys, run IDs, credentials or screen captures are published here.
- Setup diagnostics preceded the visible runs. The first automation session closed after its property reader failed. The successful session validated identities from each opened archive and again from both stored runs. A busy Excel call during the source grant and an exclusive source-hash read were resolved before destination Start. No Seed retry was issued for either visible attempt. The operator used the normal sign-in dialog. No diagnostic modal required dismissal. The disposable source and destination were retained on the review share.
- The Viewer closed through its owner. Excel quit naturally and temporary settings were restored. At 19:19 local, zero Excel processes or test controllers remained. Nine source/tooling pins, all 15 XLAM pins across the qualified, preceding and installed sets, and the five startup registrations matched. Windows Application events 1000/1001/1002 showed zero Excel faults or hangs during the review. The bootstrap template first generated during this Create was moved to the private review evidence. Unrelated working-tree changes were preserved.

## Do Not Repeat

The qualification fixture's direct Auth setup and in-package guard instrumentation are unsuitable for this visible review. Use the supported Create and Users & Roles commands and pristine candidate files. A shared saved-file read handles Excel's open source projection. Handoff 269's earlier mixed-package VBE crash remains unexplained.

## Assumptions to re-verify

Check the exact candidate and reader-reopen registration before another live review. A previously loaded Viewer publication can lag a newly saved receipt. Repeat evidence must use the repeat recording.

## Open questions and blockers

Justin later clarified that workbook windows repeatedly flashed in front of the
form and made it harder to use. A capture during guide selection showed a
read-only Config workbook visible alongside Viewer. The originating call and
frequency remain unverified. For a focused reproduction, trace window-show and
foreground changes during guide selection, Start, refresh and Verify on the
same candidate. Keep diagnostic workbook opens outside the measured clicks.
Record the affected action and workbook role before changing a loader.

Broader B1 interruption, nested entry and partial-outcome coverage remain open. Formal 4be-A, 4be-B and Release 1 acceptance are separate. The proposal below carries no implementation or recovery-policy approval.

## Immediate next action

For the proposed interruption increment, first test a lost acknowledgment after ordinary Admin Seed has applied in a fresh Training destination.

| Boundary | Existing protection or proposed assertion |
| --- | --- |
| Stop or close before dispatch | The stopped-Seed test records REQUIRED_STEP_MISSING, zero terminal sources and preserved inventory. Receiving busy-close tests defer close until the owner returns. |
| Claim and nested entry | Core claims the bound pending Seed once. Altered and duplicate bridge requests are tested. A second call after a lost response is the uncovered combination. |
| Owner applied, response lost | In a disposable Operations test copy enter through Start, count one fixed bridge call, then raise a controlled error after its real return and before consuming its result. Exercise the runner's ordinary error branch. With a valid pair, current FinishStep(False) records Failed/OWNER_NOT_COMPLETED with the APPLIED step reason and exact event. Missing capture follows the existing Unknown path. |
| Verification | Independently compare the one durable applied receipt, event lines and keys. Verify may establish the attributable saved business effect. It must preserve the failed dispatch history and reject a different run's event. |
| Reopen | Close and reopen setup without Start. The original sealed run remains readable by its own identity. There must be no further owner call, receipt or run. Starting anew is an explicit operation. |

Behavioral RED includes double dispatch, a second receipt, false dispatch completion, loss of the matched APPLIED owner evidence, or Verify using the wrong event or rewriting dispatch history. Retain the focused Seed positive, refusal, stopped and no-op controls, exact source preservation and Receiving busy-close checks. A passing first run establishes coverage without a product change.

## Critical references

Architecture v4.11 D18-REPLAY-01 items 4 and 6, Plan 022 B1, Controls v1, handoffs 269–270, `frmActionPathRun.DispatchSteps`, `modExecutionRun.PendingSeed/FinishStep/Verify`, `modAdminSeedReplay.ApplyBuiltIn`, `Slice4beAdminSeedReplay.ps1`, `Slice4beSeedReplayProof.ps1` and `Slice4beReceivingRunClose.ps1`.
