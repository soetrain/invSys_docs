# Handoff 253: live Action Paths Search proposal pending

Date: 2026-11-20 Pacific. Predecessor: handoff 252.

## Goal and release outcome

Complete 4be-A under Architecture v4.11 with responsive Action Paths Search and the remaining visible Operations/Admin comparison; human acceptance remains open.

## Current verified state

Last verified 2026-11-20: code worker `work/d18-picker-save-repro` at `5d10c4b4` is tracked-clean; preserve its pre-existing untracked `tests/tooling/reports/`. Code PR17 stays draft. Docs `work/perplexity-handoff218` at `e73f5f8` before this handoff is pushed to draft PR8. Preserve the unrelated guidance 024 deletion, handoff 067 edit and four untracked user files. Excel is closed. Handoff 251's unchanged five XLAMs remain the only review candidate; no search experiment was deployed.

## Decisions and constraints

Justin's screenshots show the accepted Viewer Events Search filtering quickly per keystroke while Action Paths Search is unusably slow. The earlier submit-on-Enter proposal was revised before approval. **Proposed, unapproved** `D18-SEARCH-INPUT-01` in Architecture v4.11, Plan 022 and Controls 1.783 now proposes live local filtering of already authorized recording headers. It explicitly changes the current per-action policy rule only for local presentation: typing checks captured session, clears selected evidence/actions, and performs no Config/NAS read; Open, Refresh, selection and other Core-backed actions recheck current policy. Previously loaded headers may remain filterable until Refresh after an external policy change; no observations or dependent actions can be obtained from them without a current-policy check. Do not implement before Justin approves this architectural exception.

## Evidence and traceability

Symptom -> path -> rule: `frmActionPaths.mSearch_Change` calls `RefreshPaths` and Core `modActionPathRead.ListPaths`, which validates Config and rereads NAS journals per character; Architecture v4.11 4be.4 currently requires policy on every library action. The exploratory packaged run was 122 PASS / 1 behavioral FAIL; two 128/128 cached builds were disqualified and their source/test edits reverted. Details and ignored report locations are in handoff 252. No user operational workbook changed.

## Do Not Repeat

Do not implement the now-superseded submit-on-Enter proposal, ship the disqualified cached builds, or count their GREEN as acceptance. Do not replace `deploy/current` or handoff 251's candidate.

## Assumptions to Re-verify

Recheck approval, branch/PR heads, worktree status, Excel closure, candidate hashes and NAS connection before building or human retest.

## Open questions and blockers

Approval of revised `D18-SEARCH-INPUT-01` is pending. Human NAS search usability, same-task How-To/Diagnostic/Compare both, incomplete-run and B0 acceptance remain open; 4be-B and final Release 1 gates are separate.

## Immediate next action

After approval, create an actual-handler native-keyboard RED for live filtering and policy-safe selection, then implement only the approved behavior and requalify a new five-XLAM candidate before Justin retests.

## Critical references

Architecture v4.11 4be.4 and proposed `D18-SEARCH-INPUT-01`; Plan 022 top checkpoint; Controls 1.783; `frmInventoryViewer.mTxtSearch_Change`; `frmActionPaths.mSearch_Change`; `modActionPathRead.ListPaths`; `Slice4beRecordingReader.ps1`; handoffs 251-252; code PR17; docs PR8.
