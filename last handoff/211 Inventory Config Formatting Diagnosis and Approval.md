# 211 Inventory Config Formatting Diagnosis and Approval

## 1. Goal and release outcome

Complete4be-A Recording/Event Viewer/How-To and required early Receiving B0 under
Architecture4.11/Plan022. **4be-A remains incomplete;4be-B/fullR1 are separate.**

## 2. Current verified state

Last verified2026-10-21 15:50Z. Implementation worktree `../invSys_edit_receipts`,
branch `work/d8-inv-edit-receipts`, is pushed `6304d89e`; draft code PR4 targets
soetrain/main. Docs main was pushed `1774d05` before this proposal revision. Main code remains
`6d247543`; preserve its user-modified AGENTS.md and unrelated docs changes
(old024 deletion, handoff067, untracked023/renamed024/025).

Excel is closed; latest controller confirms settings/package preservation and
clear delayed audit. Desktop monitor21500 stopped normally with zero sticky error5/
lock/mouse-lock counters. Catalog47/201 remains draft, unbuilt and unaccepted.
The frozen catalog46 candidate is main `deploy/validation-inventory-uom-02`.
Controls1.717 and Plan022 match the tightened proposal (documentation-only review,
2026-10-21; no new runtime verification).

## 3. Decisions and constraints

**D8-INV-CONFIG-01 is PROPOSED, not approved.** The tightened proposal switches
only interactive modal Add/worksheet ADD, including retries, to a separate Core
entry point. No actual Config change means no save and identical file bytes.
Shared/automation APIs, headless Add, catalog-only completion, Seed, retirement,
Tester, Bootstrap and migration retain formatting/save behavior; Architecture
names their exact callers. D13 adds retained-caller runtime formatting/output
regressions. Existing authority/ownership checks remain. No product code implements it.
Stop for approval before implementation. This PC retains its established origin
feature-branch/code and origin/main/docs workflow; laptop-only defaults do not apply.

Revised user stop rule: exit on lockscreen, verified mouse/input failure or required
proposal approval. Otherwise diagnose error5 and continue safe work. Confirm Excel
cleanup before another test/build. Passive desktop samples are not fresh mouse proof.

## 4. Evidence and traceability

- Prior shutdown eventually restored settings/packages; its intervened exit is
  not relabeled normal. A new unassisted shutdown also temporarily returned native
  exit0/wait258, then completed after about six minutes. This was not a lock.
- Expanded worksheet cohort is excluded: normal and failed-save/reopened-retry
  business checks pass, but the queue fault instrumented only the server writer.
  Test now covers server and local staging and checks fault reachability first;
  corrected queue RED remains unrun.
- Config diagnostic: actual Upload invokes formatting during three inbox-setup
  saves. A test-only nonformatting comparison removes worksheet/style differences.
  Remaining two table HeaderFormat references resolve to byte-identical format
  records; other table fields/tails, data and Auth remain preserved.
- Latest comparison (test commit30b13465) is21 PASS/18 FAIL: missing recording and
  strict Config classification remain failed. Controller
  `inventory-worksheet-noformat-diagnostic/red-68f0a73039494193a1620add4207b516`,
  worker `slice4be-admin-inventory-entry/b82deb2afac24cec871a349b4e7952f2`, under main
  `reports/runtime/`. Unassisted closure15:48:17Z; audit15:48:37Z clear.
- Diagnostic redaction33/33; prior cause-assessment12/12. Source is unchanged from
  draft55d89c41; its static/retained evidence keeps handoff210's scope.

## 5. Do Not Repeat

Do not force-kill Excel merely because cleanup is slow or infer closure from
HasExited/exit0. Do not kill a controller retaining restoration state. Do not count
missed fault hooks as product RED or counterfactual suppression as product GREEN.
Do not whitelist arbitrary table/style changes or substitute a newer cohort's
Config diagnosis for an older cohort's missing evidence.

## 6. Assumptions to Re-verify

Recheck Git, closed Excel and a fresh bounded cursor proof before runtime work.
The desktop monitor is stopped. Existing runtime controllers name its old report;
use a fresh report/stop file when resuming.

## 7. Open questions and blockers

D8-INV-CONFIG-01 approval is pending. No Config classifier relaxation or production
formatting fix is implemented. Recording GREEN/build, partial/multiple-event and
catalog-only EDIT coverage, guide PARTIAL/conclusion proof, modal mutation recording,
retained regressions and wider4be-A/earlyB0/live/full-chain/visible gates remain.

## 8. Immediate next action

After approval, extend actual Add/Upload RED with custom Config widths/fonts and
no-change save detection on the unchanged candidate, recording retained-caller
formatting/output baselines before implementing the separate nonformatting path.

## 9. Critical references

- Architecture4.11 D8-INV-CONFIG-01, D8-INV-01 and D18 worksheet refinement.
- Code `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`,
  Config diagnosis section: exact cohorts, test commits, exclusions and remaining work.
- Code `tests/integration/plan022_slice4be_inventory_worksheet_recording_results.md`.
- Core `modConfig.EnsureStationInbox`/`EnsureConfigSchema`/`FormatConfigSurface`;
  Admin `modAdmin.EnsureDemoStationInboxes`.
- `Slice4beInventoryWorksheetRecording.ps1`, `Slice4beInventoryCommandBoundaries.ps1`;
  `UploadNoFormat` is RED-only diagnostic instrumentation, never a product mode.
