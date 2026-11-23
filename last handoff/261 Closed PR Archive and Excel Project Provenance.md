# Handoff 261: closed draft PR archive; Excel project provenance

Date: 2026-11-23 Pacific. Predecessor: handoff 260.

## Goal and release outcome

Complete invSys Slice 4be-A under Architecture v4.11 and Plan 022. The 15 open draft PRs were closed at Justin's direction without merging; the D18-SEARCH-INPUT-02 candidate still needs a valid visible NAS review.

## Current verified state

Last verified 2026-11-23: code worker `C:\Users\Justin\repos\invSys_edit_receipts` remains on `work/d18-picker-save-repro` at `a062655a`; docs `C:\Users\Justin\repos\invSys_docs` remains on `work/perplexity-handoff218`, with the acceptance correction committed as `bdcd758` before this handoff. All 11 code and four docs PRs listed below were closed unmerged; `gh pr list --state open` returned `[]` for both soetrain repositories. No branch was deleted, no PR was merged, and `main` was not changed. Each exact PR head is retained locally under `refs/archive/closed-prs-20261123/<number>` in its corresponding repository. Preserve the pre-existing untracked `tests/tooling/reports/` and unrelated docs changes. Excel PID 29804 had zero open workbooks and quit normally; no Excel remains from that review.

## Decisions and constraints

Justin explicitly requested closure of all 15 open PRs. GitHub's connected integration returned HTTP 403 on the first close attempt; the locally authenticated `gh` CLI closed them. Closing PRs does not approve or reject their contents. Architecture v4.11 D18-SEARCH-INPUT-02 remains approved; Plan 022 and Controls 1.798 identify the November 23 visible review as invalid for the candidate. D18-RECORDING-LIST-01 remains proposed. No source, XLAM, schema or contract changed in this session.

## Evidence and traceability

- The visible forms showed unfiltered picker Search, slow Published guides Search, nonclosing forms and the old **Search name, tags or ID** label. Excel's Add-ins collection named the five `validation-search-extension-20261122` paths, but read-only VBE inspection showed all five **running** projects were from `validation-action-search-final-20261120`. The candidate form source says **Search Subject, Title, Tags or ID**. Thus the reported failure belongs to the old loaded projects; a fresh-process candidate review is still required. No workbook was open when `Application.Quit` ended Excel naturally.
- Preserved code PR chain: #4 `550d9ff0` -> #11 `0e014dc8` -> #12 `c4444f58` -> #13 `f9016668` -> #14 `b8f94241` -> #15 `0e988b2c` -> #16 `6e866f16` -> #17 `a062655a`. Exact Git ancestry confirms these heads are contained in #17. Separate, unmerged code heads: #5 `97716fa2`, #7 `62b3da4d`, #8 `7d78b9cf`; none is contained in #17 or `main`.
- Preserved docs chain: #3 `5557cae9` -> #5 `e8880ebd` -> #8 `916a78d7` at closure. Exact Git ancestry confirms #3 and #5 are contained in #8. Separate, unmerged docs #7 `5e089c22` is in neither #8 nor `main`. Full SHAs remain available through the local archive refs and closed PR URLs. The docs branch subsequently gained the provenance correction and this handoff; its remote branch remains available without an open PR.
- The November 22 candidate's focused packaged and native results remain as recorded in handoff 260. They are technical evidence on small fixtures, not visible NAS acceptance.

## Do Not Repeat

Do not infer running VBA projects from Excel's Add-ins registration. Require exact live VBE project paths after a fresh Excel start. Do not count the November 23 old-project screenshots as candidate RED or claim the candidate passes the larger-list NAS review.

## Assumptions to Re-verify

Check the code/docs branch heads, local archive refs, candidate XLAM hashes and Excel closure before resuming. Closed PR head branches may be moved or deleted later; the local archive refs and closed PR pages preserve the captured heads. Justin must enter invSys credentials privately for the next visible review.

## Open questions and blockers

Search actions, Published guides Search and Close/X have not yet been tried in a verified fresh process running the November 22 candidate. Slice 4be-A, B0 and Release 1 acceptance remain open. If the same defect reproduces under verified candidate paths, capture the actual handler state and perform D13 RED before changing code.

## Immediate next action

Start Excel fresh with the five `validation-search-extension-20261122` XLAMs, verify all five live VBE project paths match that directory before opening forms, then obtain Justin's disposable WHT7025AE Search and Close/X result.

## Critical references

Architecture v4.11 D18-SEARCH-INPUT-02; Plan 022; Controls 1.798; `frmGuideActionPicker.mSearch_Change`, `frmActionPathLibrary.mSearch_Change`, handoff 260, closed code PRs #4/#5/#7/#8/#11-#17 and docs PRs #3/#5/#7/#8.
