# Handoff 245: disposable Training Action Path preview and Excel crash

Date: 2026-11-15 Pacific. Predecessor: handoff 244.

## Goal and release outcome

Reach invSys Release 1 and Slice 4be-A human acceptance under Architecture v4.11 D18 and Plan 022. The first disposable Training preview exposed a guide interaction/native crash; no Action Path, B0 or Release 1 acceptance is claimed.

## Current verified state

Last verified 2026-11-15 Pacific: code `main` is `6d247543` with Justin's unrelated modified `AGENTS.md`; draft code PR16 remains `work/d14-packaged-chain-coverage` at `6e866f16`. Docs `work/perplexity-handoff218` was based on `2813480` before this handoff, with unrelated deletion/edit/untracked files preserved. The current five candidate XLAMs are registered for Excel startup; their hashes remain recorded in ignored review `candidate.json` (PR4-BASE-18). Excel PID 30248 crashed. A replacement Excel opened AutoRecovered review-host and Receiving workbooks. Justin chose to leave them for inspection; at wrap-up (2026-11-16 07:22 UTC) no Excel process remained and no new Excel fault was recorded. Justin later confirmed he did not save any open Excel files. The original review host remains unchanged since before the review; no guide version exists. Whether any separate AutoRecover artifact remains is unknown. Do not overwrite recovery artifacts; verify Excel closure, startup registration and candidate hashes before a new test. No product code or XLAM changed this session.

## Decisions and constraints

Justin designated existing NAS warehouse WHT7025AE disposable Training. The previously cancelled D8-TRAINING-ONBOARD-01 first-use proposal stays cancelled. An authorized test-only `ACTION_PATH_MAINT` grant and privately entered sign-in enabled preview; this does not qualify normal Users & Roles grant or onboarding. Tracking policy version 1 enabled recorded-control capture. Candidate registration was intentional for review; restore normal stable startup registration only after the recovered session is safely resolved. Never discard recovered unsaved staging by assumption.

## Evidence and traceability

The visible Events Viewer loaded 23/23 available groups and 26 lines; Inventory and Shipping Holds source coverage remained unavailable. The stopped Receiving recording journal reached version 12 with five distinct controls and ten requested/result observations; its expected conclusion was None. The Viewer displayed four controls from an Events publication made before the final Confirm Writes result and Receiving Close; a later republish was blocked by Excel rejecting COM calls. Justin reported **Choose tracked actions** became unresponsive after clicking around, then an Action Path guide Save was attempted. He cannot identify which Create guide entry route was used. The warehouse guide store contains zero guide versions, so Save is unqualified. Windows Application Error 1000 at 2026-11-16 06:41:04 UTC records Excel 16.0.20527.20196 faulting in `VBE7.DLL`, `0xc0000005`; causation is unproven. The native Visual Basic dialog reported `-2147418105 (80010007)` after the server disappeared. Sanitized local evidence: ignored `reports/runtime/4be-a-human-review-20261115/{recording-vs-events-evidence.json,excel-crash-observation.json}`. Symptom -> possible picker interaction, form lifetime or guide Save fault (unproved) -> D18 direct curation and immutable guide save -> focused packaged handler test still required. Existing guide curation covers simple Cancel/Save, not this selected-recording/picker-to-Save sequence.

## Do Not Repeat

Do not attribute the crash to Cancel, the window X, the Save handler or our rejected COM calls without focused evidence. Do not count five observed controls as a verified guide conclusion, four displayed rows as complete current publication, test-only Auth setup as user-side acceptance, or an absent guide file as a successful Save. Do not treat Excel Recent as proof of Save; Justin confirmed no open file was saved. Do not overwrite any residual AutoRecover artifact.

## Assumptions to Re-verify

Recheck Excel ownership, recovery state, startup add-in paths, candidate hashes, branch/PR heads and NAS target before any runtime action. The exact clicked picker controls, guide editor entry route, and whether a residual Receiving AutoRecover copy remains are unknown. The Event Viewer may still hold the earlier loaded publication. Native crash causation is unknown.

## Open questions and blockers

Excel is closed and Justin confirmed no manual save. Any residual AutoRecover copy must remain untouched; an Excel-free test lane is available after startup/package checks. Both guide entry routes require reproduction because the operator cannot distinguish them. Inventory and Shipping Holds publication availability remain separate 4be-A coverage gaps. The ordinary `ACTION_PATH_MAINT` grant and first-use sign-in gap remain open.

## Immediate next action

After confirming Excel remains closed and candidate pins are unchanged, create and run a focused packaged RED through both real Create guide routes, picker interaction and Save in a selected stopped-recording context, checking responsive close, immutable guide persistence and natural Excel exit before changing product code.

## Critical references

Architecture v4.11 4be.5/D18; Plan 022; Controls 1.760; `frmGuideActionPicker`, `frmActionPaths`, `frmActionPathGuide`, `modGuideEditor`, `modCuratedGuideSource`, `modActionGuideDraft`, `tests/tooling/Slice4beGuideActionCuration.ps1`, `tests/tooling/Slice4beGuideDraft.ps1`; `tools/register_current_addins.ps1`; ignored review `candidate.json` and `excel-registration-before-review.json`.
