# Settings diagnostics checkpoint and desktop blocker

## 1. Goal and release outcome

Complete Release 1 and Slice 4be under Architecture v4.11 / Plan 022: comprehensive
Operations/Admin Events and How-To, Diagnostic and Compare Action Paths. The goal
remains active and incomplete. Accepted focused evidence is not full acceptance.

## 2. Current verified state

- Last verified 2026-09-22: code `main` **e986b51**, docs `main` **2fb413a**, both
  pushed before this handoff. Controls v1.221. This handoff is a separate commit.
- Frozen isolated candidate: `deploy/validation-settings-diagnostic`. Runtime
  and package pins are in `reports/runtime/settings-diagnostic-runtime-pins.json`
  (299 files) and `settings-diagnostic-package-pins.json` (five packages).
- Focused diagnostics **780/780**, retaining 450 observations after 27 expected
  failures. Its shutdown was not normal; do not claim clean focused closure.
- Separate Settings **202/202** (10 accepted images), owner commands **460/460**
  (23), and Admin UOM **228/228** (11) preserve prior GREEN identities, unchanged
  candidate hashes, immediate unassisted Excel closure and zero Application errors.
- Full-chain retries pass **32/32**, live-role **48/48**, Create Warehouse
  **15/15**, but strict acceptance fails on Excel shutdown Application events.
- Broader Boxing/Shipping stops at **1,681 PASS / 8 FAIL**, 26 prior checks
  unreached. Seven failures are the existing pending D8-A cases; one is a capture
  harness exception. Nineteen of twenty images accepted; one is misframed.
- Harness DPI correction: focused geometry RED 3/4, GREEN **4/4**. This is
  harness evidence only; visible packaged rerun remains required.
- At 21:32 UTC Excel was closed, but `OpenInputDesktop(0,false,1)` returned
  NULL/error 5. User reported technical difficulties resolved; actual desktop
  access remains unverified. A follow-up request to unlock/connect is pending.
- Preserve unrelated code `src/Operations/Forms/frmEventDetail.frm` (+18 measurement
  lines, Locked=True), modified docs handoff067, and untracked critique023.
  They were excluded from both checkpoints. No operational/NAS deployment changed.

## 3. Decisions and constraints

D18 semantic inheritance is approved. **D8-A and Event Detail Locked=False are
not approved.** Do not infer approval from a request to resume work. Do not change
those contracts, carrier authority or training-transfer semantics silently.

Preserve D5/D12/D13, exact immutable System_Key, unknown columns, captured workbook
and session, packaged launcher reuse, headless authority and all previous GREENs.
Run Excel gates serially; never rebuild/deploy while relevant Excel is open.
Use disposable fixtures and the frozen candidate; do not touch operational files.

## 4. Evidence and traceability

The committed detailed record is
`tests/integration/plan022_slice4be_settings_editor_activity_results.md` in the code
repository. It records the failed attempts, RED/GREEN, exact report roots and
limitations. Plan 022 and controls are synchronized with that checkpoint.

Ignored receipts in `reports/runtime/`:

- `settings-diagnostic-focused-behavior-verification.json`: 780 behavioral checks;
  no clean-shutdown acceptance.
- `settings-diagnostic-compiled-settings-verification.json`: 202 checks.
- `settings-diagnostic-capture-regression-owner-verification.json`: 460 checks.
- `settings-diagnostic-uom-verification.json`: 228 checks.
- `settings-diagnostic-regression-boxing-attempt-verification.json`: incomplete
  1689-check attempt, preserved pins and normal closure.
- `settings-diagnostic-chain-retry-behavior-verification.json` and
  `settings-diagnostic-chain-trace-behavior-verification.json`: passing behavior,
  two Application events each, clean/full release acceptance false.
- `settings-diagnostic-static-verification.json`: 250 components, 6,045 procedures,
  132,893 lines, 9 literal / 45 unresolved Application.Run calls, 193 duplicate
  groups. The one-group increase is a documented bounded exception for 26 actual
  MSForms event adapters. Three schemas and all 28 existing size ratchets pass.

The capture defect is independently proven: virtualized GetWindowRect bounds
were used with physical CopyFromScreen pixels. Scoped thread DPI awareness fixes
the focused dimension mismatch and restores caller context. This does not prove
that error 5/cursor-access failures are fixed.

## 5. Do Not Repeat

- Do not label compile, capture, unavailable desktop or harness failures product RED.
- Do not repeat a used runner prefix or overwrite failed evidence.
- Do not infer process termination from HasExited/ExitCode alone: native wait was
  previously unsignaled despite reported exit. Do not bypass the Excel guard.
- After a traced chain stage's original Quit returns, use native process/window
  metadata only. Do not externally reattach COM/GetActiveObject or issue extra
  Quit. Prior AVs correlated with such interaction; causation is unproven.
- Recovery choices require viewing a rendered image, then a separate action.
  One traced run's unpainted retention image was incorrectly followed by Confirm;
  its corrected receipt explicitly does not verify recovered-file retention.
- Do not retry the earlier rejected deletion of the temporary directory ending
  `invsys-config-command-c365fc6fbfab48a79b8f842127ebf3fd`; leave it intact.

## 6. Assumptions to re-verify

Input desktop access, Excel closure, Git state and package hashes. Test hashes
recorded during completed gates covered 188 files; the new geometry test makes
future test pins 189. Do not rerun old verifiers against subsequently edited test
files and reinterpret the original receipts. Runtime/package pins remain fixed.

## 7. Open questions and blockers

- Restore accessible input desktop before visible tests. Pending user question is
  about unlocking/connecting the desktop, not permission to continue authorized work.
- D8-A approval: seven `Shipping.Access.AuthUnavailable.*.MissingFileNotRecreated`
  failures remain real failures, never GREEN. Event Detail scrolling approval also
  remains pending; preserve user edits and Locked=True.
- D5 carrier conflict: Core `modCarrierSettings` reads/writes Windows-user registry
  settings, while D5 specifies warehouse Config authority. Admin carrier handlers
  have no explicit ADMIN_MAINT check. Reconcile before defining new owner metadata
  or changing persistence; do not silently copy UOM semantics.
- Connection option is distinct: Admin handler checks cached ADMIN_MAINT, but Core
  `modNasConnection.SetRequireManualServerCredentials` saves a Windows-user preference
  with no corresponding command-boundary check. Its scope/authority needs explicit
  review before new tracking outcomes; this is source discovery, not behavioral RED.
- Remaining comprehensive coverage, guide transfer design, clean release shutdown,
  physical NAS/multi-station checks and human acceptance are still open.

## 8. Immediate next action

Once input-desktop access is verified and Excel is closed, run
`tests/tooling/Test-Slice4beCaptureForeground.ps1 -RepoRoot .`, require all three
cases Captured, directly review their images and verify normal closure before
resuming the packaged gates.

## 9. Critical references

- Next unused broader gate:
  `reports/runtime/run-settings-diagnostic-capture-regression.ps1 -Gate boxing`.
  Preserve all 1,714 baseline identities, with only the seven existing D8 failures;
  require all 22 images accepted. Clone the original regression verifier with the
  new capture-prefix path only; do not relax checks.
- Shutdown control, prepared but **not run**:
  `reports/runtime/run-settings-diagnostic-chain-passive.ps1` and
  `verify-settings-diagnostic-chain-passive.ps1`. Frozen instrumented copy, original
  chain unchanged; observe only native metadata after original Quit.
- Comparison, **not run on this candidate**:
  `reports/runtime/run-settings-diagnostic-regression.ps1 -Gate comparison`.
  Retain 421 checks and 31 images, including the strengthened actual-handler
  `Recording.PolicyChangeClosesIncomplete` assertion (not part of Boxing gate).
- Other remaining current-candidate gates: curation77, restart27, Viewer94,
  guide351. Use corresponding unused runner prefixes and compare exact baselines.
- `reports/runtime/admin-settings-remaining-coverage-notes.md` and
  `reports/runtime/guide-transfer-design-notes.md` retain discovery/proposals,
  not approved contracts.
