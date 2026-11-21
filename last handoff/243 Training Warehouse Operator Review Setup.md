# Handoff 243: Training warehouse operator review setup

Date: 2026-11-15 Pacific. Predecessor: handoff 242.

## Goal and release outcome

Prepare Justin's hands-on 4be-A How-To/Diagnostic comparison under Architecture v4.11 D18-REPLAY-01 and Plan 022. Candidate and review sheet are ready; permission setup and human review remain open.

## Current verified state

Last verified 2026-11-15: draft code PR16 `work/d14-packaged-chain-coverage` carries the review sheet at `6e866f16210ac864a891a3b425e2d99001b19d1a` and remains draft. Draft docs PR5 branch `work/perplexity-handoff218` contains this handoff, Plan 022 and Controls 1.758; the pre-existing unrelated docs deletion, edit and untracked files remain untouched. Main code checkout retains Justin's unrelated `AGENTS.md` edit. Justin authorized closing Excel; the Book1 process closed without intervention. Excel is closed. Do not load packages until review setup is agreed.

The ignored `reports/runtime/4be-a-human-review-20261115/` contains a five-XLAM review copy and the sibling inventory template, all hash-matched to the unchanged current validation files. `TrainingHub` is empty. No Training warehouse or operational workbook was created or changed.

## Decisions and constraints

Justin chose **Create new Training warehouse**. D8-CREATE-01 requires a captured, signed-in source warehouse Admin session with `ADMIN_MAINT` before interactive Create; the destination has no Auth store yet. No source sign-in, credential or warehouse target has been supplied. The review sheet uses ordinary Receiving and Admin preference controls, separate guide-source and observed runs, and complete/incomplete comparisons. It records no human acceptance. Executable Run How-To remains separate.

## Evidence and traceability

The five copied XLAM hashes match PR4-BASE-18; the copied template matches the validation source. The candidate review sheet is `tests/integration/plan022_slice4be_operator_review.md` in code PR16. Read-only source inspection: `modWarehouseBootstrap` grants `ADMIN_MAINT` and `RECEIVE_POST` but `EnsureUserRow` leaves `PinHash` empty; `modAuth.ValidateUserCredentialForCapability` rejects an empty hash. `frmCreateDeleteUser` and `modAuthUserCommands` expose a six-choice mask without `ACTION_PATH_MAINT`; guide creation checks that separate capability in Core. Packaged guide tests use an explicit disposable Auth fixture grant. Symptom: new Training Admin cannot sign in or author a guide through ordinary UI -> missing user-side onboarding path -> D18's capability-gated authoring and D8's headless write authority. No behavioral RED/GREEN for a correction has begun.

## Do Not Repeat

Do not infer `ACTION_PATH_MAINT` from `ADMIN_MAINT`, treat a blank PIN as usable, edit a live Auth workbook, or count fixture setup as user-side onboarding acceptance. Do not treat copied XLAMs or the empty hub as a completed warehouse.

## Assumptions to Re-verify

Recheck Excel ownership, candidate hashes, current PR heads and empty Training hub before setup. Confirm a source Admin sign-in before Create. No architecture amendment for the grant path has been approved.

## Open questions and blockers

Justin authorized closing Excel but did not answer the grant question. A follow-up asks whether both a privately entered PIN and `ACTION_PATH_MAINT` may be provisioned only in the disposable Training fixture for an early preview, leaving ordinary onboarding open; the answer is pending. Another pending question asks whether an existing signed-in source Admin warehouse is available for the D8-CREATE-01 Create action. A proposed user-side resolution is saved locally as `reports/runtime/4be-a-human-review-20261115/D8-TRAINING-ONBOARD-01.PROPOSAL.md`; it is not approved or normative. A Release 1 onboarding path needs an architecture decision before implementation, then Plan/Controls synchronization and D13 packaged RED/GREEN.

## Immediate next action

After Justin chooses the preview setup and identifies a source Admin warehouse, establish a captured source Admin sign-in and create the new Training warehouse through the packaged Admin Create form; keep credential and guide-right setup clearly fixture-scoped according to that choice.

## Critical references

Architecture v4.11 D8-CREATE-01/D18-REPLAY-01; Plan 022; Controls 1.758; code PR16/docs PR5; `modAdminCreateEntry.OpenFromRibbon`, `modWarehouseBootstrap.BootstrapWarehouseLocal`, `modAuth`, `frmCreateDeleteUser`, `modAuthUserCommands`, `modGuideDraftSource`; local review `candidate.json` and operator review sheet.
