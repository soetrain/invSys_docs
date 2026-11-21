# Handoff 251: guide input current-candidate requalification

Date: 2026-11-20 Pacific. Predecessor: handoff 250.

## Goal and release outcome

Complete Slice 4be-A under Architecture v4.11 D18-GUIDE-INPUT-01 and Plan 022 without Release 1 regressions. The approved buffered guide input has packaged GREEN and broad current-candidate technical requalification; 4be-A human and remaining final-candidate gates are open.

## Current verified state

Last verified 2026-11-20: code worker `C:\Users\Justin\repos\invSys_edit_receipts`, `work/d18-picker-save-repro` at `5d10c4b4`, pushed to draft [code PR17](https://github.com/soetrain/invSys_fork/pull/17) against `work/d14-packaged-chain-coverage`. It has no tracked changes; preserve its pre-existing untracked `tests/tooling/reports/`. Docs `C:\Users\Justin\repos\invSys_docs`, `work/perplexity-handoff218` at `8c9da50` before this handoff, pushed to draft [docs PR8](https://github.com/soetrain/invSys_docs/pull/8) against `work/excel-shutdown-timing`. Preserve unrelated deletion of guidance 024, edit to handoff 067, and the four unrelated untracked files shown by `git status`. Main code checkout remains at `6d247543` with Justin's `AGENTS.md` edit.

The unchanged candidate is the five XLAMs in `reports/runtime/guide-input-green6-20261117/addins/` under the worker. Operations SHA256 `6FAF34132B35B3A10FBA95F9A6D01B7AB20E2DF572C3BE22ED47740679C4AFE3`; Core SHA256 `43DF458CC4A6533839567ACC810F552A96ABD3CCCC815D2DD783D415C7BDBC01`. Excel count was zero at the final audit. Isolated chain worktree `C:\Users\Justin\repos\invSys_chain_20261120` stays detached at `17675c6c`; its generated `tests/integration/slice14_results.md` has a tracked modification with synthetic row details. Keep local and sanitize before publication. No candidate deployment to `deploy/current` or NAS was made.

## Decisions and constraints

`D18-GUIDE-INPUT-01` is approved. Name, Tags, Instructions and Step Instruction buffer locally; dirty text retains its StepId; dependent actions/Save/Tab/mouse transitions flush through typed Core, while dirty Cancel/X discards without a write. No contract change was made this session. Test-only transfer-policy fixture commits `f7bcd4bf` and `5d10c4b4` use Core's exact catalog-29 IDs and an actual-handler-authored older-safe guide; product XLAMs are unchanged. Controls is now 1.781 and Plan 022 has the matching checkpoints.

## Evidence and traceability

- Focused guide input and preservation: packaged **171/171**, input-order **56/56**, evaluation/How-To/Diagnostic/Compare **247/247**, five compiles each. Prior 169/171 was a stale test assertion after legitimate activity additions, corrected test-first.
- Current candidate: guide transfer **68/68**, recorded transfer/diagnostic **496/496**, policy/activity **426/426** (older policy, user/control off, store failure and recovery); preference saved-copy restart **28/28**; Event Detail **81/81**; visible Settings **202/202**; Viewer groups/Admin publication **86/86**; Viewer published-read/recording **80/80**. Reports are in the corresponding ignored `reports/runtime/slice4be-*` families; the policy controller receipt is `reports/runtime/guide-transfer-controller/b0a77f686bd64cbb99d938fa3f68a228/closure.json` (exit 0, settings restored, packages preserved, Excel closed). The 496-check session took about nine minutes after its GREEN line to exit naturally, with no owned modal or Excel Application fault found; slow exit remains unresolved.
- Current candidate early Receiving B0 technical replay **121/121**, all four scope flags true and ten critical identities passed, but the harness deliberately records `B0Accepted=false`; human B0 acceptance is open. Current-candidate isolated Release 1 chain **33/33** and standalone live-role **48/48** passed with natural Excel closure and zero checked Application faults. Static maintenance: 391 components, 6760 procedures, 9 literal/45 unresolved dynamic calls, 194 duplicate-body groups; only three new reviewed MSForms handler groups, six reviewed and zero unreviewed warning deltas. See Plan 022 checkpoints and local ignored reports for exact receipts.
- Operator review in the disposable NAS warehouse `WHT7025AE` remains incomplete. A byte-identical review-host copy was prepared at `reports/runtime/4be-a-human-review-20261120/ActionPath-Review-Host.xlsm` in the worker; it has not been opened for this review. The operator availability question sent this session was unanswered at handoff time.

## Do Not Repeat

Do not treat a catalog-47 policy relabelled as catalog 29 while retaining later controls as valid. Do not use the current six-action guide under catalog 29: later navigation controls become hidden and Export/Import success expectations are invalid. The resolved fixture creates a one-action catalog-29-safe guide through actual handlers. Do not combine standalone gate totals or equate technical B0 replay with `B0Accepted`.

## Assumptions to Re-verify

Recheck Excel closure, candidate hashes, active input desktop if native interaction is needed, NAS access, source and docs PR heads, and both worktree statuses. Excel build at the last Settings run was 16.0.20527.20196; updates can change it. The operator may prefer another review time.

## Open questions and blockers

4be-A remains open under Architecture v4.11: visible operator comparison of How-To and Diagnostic for the same Operations and Admin tasks including an incomplete run, NAS list/step usability, human B0 acceptance, and the remaining final-candidate Inventory, launcher, reusable Production and Release 1 acceptance gates. The policy gate is GREEN but does not close these. The underlying intermittent slow Excel native exit remains unresolved; the Microsoft report has already been submitted. 4be-B replay and Admin result management come later.

## Immediate next action

With the operator present, open the prepared disposable NAS review host on the unchanged candidate and record visible How-To/Diagnostic/Compare both behavior for one Operations and one Admin task, including an incomplete run and NAS list responsiveness; keep human acceptance separate from packaged results.

## Critical references

Architecture v4.11 `D18-GUIDE-INPUT-01` and 4be-A completion rule; Plan 022 top checkpoints; Controls 1.781; code PR17; docs PR8; `frmActionPathGuide.frm`; `Slice4beGuideTransferPolicy.ps1`; `Test-Slice4beConfigCommands.ps1`; the candidate addin folder and ignored receipts named above.
