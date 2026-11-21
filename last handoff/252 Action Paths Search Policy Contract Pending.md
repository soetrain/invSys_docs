# Handoff 252: Action Paths Search policy contract pending

Date: 2026-11-20 Pacific. Predecessor: handoff 251.

## Goal and release outcome

Complete 4be-A under Architecture v4.11 with usable NAS Action Paths Search and the remaining visible Operations/Admin comparison. Human acceptance remains open.

## Current verified state

Last verified 2026-11-20: code worker `work/d18-picker-save-repro` at `5d10c4b4`, with no tracked changes and pre-existing untracked `tests/tooling/reports/`. Code PR17 remains draft. Docs `work/perplexity-handoff218` at `0c9857e` before this handoff; proposal is published to draft docs PR8. Preserve the unrelated guidance 024 deletion, handoff 067 edit and four untracked user files. Excel is closed. The five-XLAM candidate from handoff 251 remains unchanged and is the only review candidate; the review host is saved in ignored `reports/runtime/4be-a-human-review-20261120/`.

## Decisions and constraints

Justin observed unusably slow typing in **Action Paths > Search recordings** during NAS operator review. Architecture v4.11 4be.4 requires every library action to check captured context and current policy. A local snapshot filter on every keystroke conflicts with that rule. Experimental source and test edits were reverted; no experimental XLAM was deployed. **Proposed, unapproved** `D18-SEARCH-INPUT-01` in Architecture v4.11, Plan 022 and Controls 1.782 treats typing as pending local input; Enter or the existing Refresh button executes a current-policy Core Search. Do not implement until Justin approves it. No new control or ControlId is proposed.

## Evidence and traceability

Symptom -> path -> rule: `frmActionPaths.mSearch_Change` invokes `RefreshPaths`, which calls Core `modActionPathRead.ListPaths`; `modTrainingReadContext.Read` validates policy through `modActivityPolicy.ReadPolicy` and the reader enumerates NAS journals and first entries. NAS enumeration alone took about 400 ms. Focused packaged experimental test: **122 PASS / 1 behavioral FAIL** (`RecordingRead.SearchUsesLoadedSnapshot`) on the prior five XLAMs; no fixture or compile failure. Two experimental cached builds reached **128/128** with five compiles and natural closure, but neither qualifies: one still checked Config per key, and the final build skipped the current-policy check for Search. Final experimental static scan: 391 components, 6761 procedures; it is not candidate evidence. All experimental packages remain only in ignored `reports/runtime/4be-action-path-search*-20261120/`. Review Excel closed naturally; no operational workbook changed.

## Do Not Repeat

Do not ship the local-snapshot `mSearch_Change` experiment or count its 128/128 as 4be-A GREEN. It violates the operative every-library-action policy rule or retains per-key Config cost. Do not overwrite the handoff-251 candidate or `deploy/current`.

## Assumptions to Re-verify

Recheck approval, both repository statuses, PR heads, Excel closure and the five candidate hashes before a new build. The operator's earlier test question referred to a now-discarded validation build; disregard any result from it.

## Open questions and blockers

Approval of `D18-SEARCH-INPUT-01` is pending. Human NAS typing, same-task How-To/Diagnostic/Compare both, incomplete-run and B0 acceptance remain open. 4be-B and final Release 1 gates remain separate.

## Immediate next action

After approval, add a native-keyboard packaged RED for pending typing and Enter/Refresh policy checks, then implement only the approved submit-on-Enter-or-Refresh behavior and requalify on a new five-XLAM build.

## Critical references

Architecture v4.11 4be.4 and proposed `D18-SEARCH-INPUT-01`; Plan 022 top checkpoint; Controls 1.782; `src/Operations/Forms/frmActionPaths.frm`; `src/Core/Modules/modActionPathRead.bas`; `tests/tooling/Slice4beRecordingReader.ps1`; handoff 251; code PR17; docs PR8.
