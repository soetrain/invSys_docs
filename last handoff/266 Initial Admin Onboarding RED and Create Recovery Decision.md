# 266 Initial Admin onboarding RED and Create recovery decision

## Goal and release outcome

Qualify an ordinary authorized Create, destination Sign In, explicit scoped grant, guide Author/Edit/Transfer, Revoke and cold reopen under D8-CREATE-01, D8-USERS-01 and D18. Justin approved this increment on November 24, 2026 against docs `52184a4`. The packaged Create baseline reached the approval's explicit recovery-policy stop before product implementation.

## Current verified state

Last verified November 24, 2026 Pacific. Code `work/d18-picker-save-repro` advanced from `1db8bbfa` to `ecb497aa` (test files only). Docs `work/perplexity-handoff218` publishes to `work/d18-picker-save-repro-docs`. Approval sync `b7d33bc` updated Architecture v4.11, Plan 022 and Controls v1. This handoff, its pointer and the Plan/Controls baseline notes follow that commit. Prior code PR17 and docs PR8 remain closed. The primary code checkout's AGENTS edit, the worker's untracked tooling reports and unrelated docs edits are preserved. Excel closed naturally after RED.

## Decisions and constraints

Keep the existing source authorization, exact account/scope, two bootstrap roles and credential format. A new recovery policy requires Justin's decision. The existing reopened Publish command reads an already-created warehouse and may not re-bootstrap it or repair credentials. Tester Setup/Delete ownership receipts apply to that separate command. There is no corresponding receipt for the ordinary Create run tested here.

## Evidence and traceability

- Code `tests/tooling/Test-Slice4beConfigCommands.ps1` exposes the focused selector, `tests/tooling/Slice4beAdminCreateWarehouseOwner.ps1` connects it to the real Admin Create handler, and `tests/tooling/Slice4beInitialAdminOnboarding.ps1` checks blank entry and injects a one-shot refusal at the existing destination account provisioning boundary. Source fixture setup supplies an authorized creator. The test makes no direct destination Auth edits.
- Inputs were the five catalog 48 packages at `deploy/validation-guide-reader-reopen-20261124`. All five instrumented projects compiled. All five original SHA-256 hashes matched `reports/runtime/4bea-picker-finish-20261123/buildReaderReopen/package-pins.json` after the run. Baseline `reports/runtime/slice4be-admin-create-owner/2ef27a72b12d462f93f5c1f6b313b31b/red.json`: 10 passed, 3 failed. The runtime log is `reports/runtime/onboarding-20261124-baseline.log`. No modal intervention or forced termination. Zero remaining Excel processes and zero matching Excel Application Error, Hang or WER events in the run-window audit.
- Failed checks: `Onboarding.EmptyPin.RefusesBeforeBootstrap`, `Onboarding.EmptyPin.NoDestination`, `Onboarding.ExistingHub.IncompleteDestinationNotPublishable`. Blank entry reached actual bootstrap and created a Training destination. The account-provisioning fault was consumed in both folder scenarios. Both returned creation failure. A run-created root was removed. In an existing root the unrelated file's hash survived, Config and Auth remained, and the read-only validator used by Publish Existing accepted the partial warehouse. A second actual Create click did not reach bootstrap. No share publication was attempted.
- Trace: `modAuth.EnsureStationRoleAuth` opens/creates destination Auth before the account write. `modWarehouseBootstrap.BootstrapWarehouseLocal` removes the failed root only when that call created it. Its duplicate guard rejects surviving runtime files. `modWarehousePublishSource.ReadExisting` accepts the saved Config/station plus Auth, Inventory and inbox presence. It cannot distinguish this failure from a completed hub.
- No corrected GREEN, destination credential-writer test, ordinary destination sign-in, grant/author/revoke or cold-reopen result. Exact-scope/stale authority, six-choice preservation, already-open-form revocation, guarded flush, dirty Cancel/X, and new-field privacy are unqualified. No destination PIN was entered or recorded in this baseline. Layout, private input, regression, fresh-build qualification and visible operator acceptance remain pending. Historical guide/B0/Release 1 results retain only their original scope.

## Do Not Repeat

Do not count eligibility as actual publication, treat an account-provisioning fault as proof of the future PIN writer, use a prepared destination account for qualification, or delete a pre-existing directory to make retry pass.

## Assumptions to Re-verify

Recheck source/docs heads, package pins, startup bindings and Excel ownership before resuming. Keep catalog-49 and picker feedback work separate.

## Open questions and blockers

Decision requested: may Create roll back only files and empty directories it can prove it created during a failed initial Admin credential write inside an existing hub folder? Preserve the original directory and all pre-existing content. Where ownership or cleanup is uncertain, report failure and keep the incomplete hub ineligible for publication. This is a proposal. No partial-hub deletion, quarantine or recovery mechanism was added.

## Immediate next action

After Justin selects a recovery rule, record it in D8-CREATE-01 and run a focused packaged test of that exact rule, including preservation, publish refusal and the agreed retry behavior.

## Critical references

`0 plan docs/xlam_invSys/invSys-Design-v4.11.md` (D8-CREATE-01); `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`; `0 plan docs/xlam_invSys/invSys-Controls-v1.md`; `frmCreateWarehouse.btnCreate_Click`; `modWarehouseCreateCommands.CreateWarehouse` and `PublishExisting`; `-AdminCreateWarehouseOwnerOnly -CheckInitialAdminOnboarding -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 -Phase RED`.
