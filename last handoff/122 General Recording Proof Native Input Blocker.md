# General recording proof: native input blocker

## 1. Goal and release outcome

Complete Slice4be-A under Architecture v4.11/Plan022; active acceptance work is
General Settings independent recordings and guide-conclusion comparisons.

## 2. Current verified state

Last verified:2026-10-08,21:48 UTC (14:48 PDT).

- Code main `7b93e961`, pushed and clean. New tests are staged in source, not accepted.
- Docs main before this ending update: `58d5ecb`, pushed; this update synchronizes
  Plan022 and Controls1.500 without changing the normative contract.
- Runtime remains `6f583238`; five frozen packages in
  `deploy/validation-general-settings-02` are unchanged. Retain General647/UOM228.
- First recorded diagnostic:17 passes/1 native harness failure before new cases;
  all five instrumented compiles pass. Settings/packages preserved, unassisted
  Excel closure, delayed Application audit zero failures. No Excel remains open.
- Desktop monitor has exited. No recurring helper was installed or left running.
- Preserve user-owned handoff067 edit, old024 deletion and untracked023,
  REPORT024 and guidance025; their bytes were rechecked unchanged.

## 3. Decisions and constraints

The user corrected handoff121's interpretation: the prior14:31 error5 followed
RDP X-close, not script disconnection. They now report script disconnection worked.
This run had **no error5**, but native pointer input failed. Do not conflate these.
Stop and timestamp any later actual error5; preserve test-owned cleanup and exit.
No Windows/session settings, native capture helper or product runtime were changed.
D18-REPLAY-01 remains approved; Scale/UI decisions remain separately pending.

## 4. Evidence and traceability

- Full receipts and test scope:
  `tests/integration/plan022_slice4be_general_settings_results.md`.
- Trial controller `reports/runtime/general-settings-controller/85030cc2c0ed4d3bb66fb9059b0394d4`;
  worker `slice4be-general-settings/858abe276ffc4a3baca9de01ca400961`.
  `settings-save.png` baseline focus failed in `ActivateByCaptionClick` before
  General recording cases. This is not product RED or new GREEN.
- Read-only monitor: `reports/runtime/rdp-desktop-probe-cae9576a9b1d409e9dc15cf39ca1498d/summary.json`;
  21:33:51--21:48:12 UTC (14:33:51--14:48:12 PDT),170 samples, zero cursor-read,
  input-desktop, pixel-capture or error5 failures. Exact disconnect time unverified.
- Ignored `tscon-cursor-permission-observation.json` and
  `tscon-cursor-after-console-transition.json`: SetCursorPos=false/error0.
  `tscon-desktop-identity-observation.json`: thread/input desktops both Default.
  `tscon-owned-native-input.json`: SendInput returned1 but cursor did not reach
  the verified disposable target; no click sent. Cause remains unresolved.
- Existing capture geometry passes4/4; it does not prove pointer input.
  Static `general-settings-recorded-static-01/ratchet-verification.json` retains
  319 components/6340 procedures/138712 lines,9/45 dynamic calls,192 duplicate
  groups and28 oversized caps; three schemas validate,451 scripts parse.

## 5. Do Not Repeat

Do not retry the full Excel gate while native input is unavailable, bypass visible
evidence, change Windows policy, add keep-awake, or claim successful desktop reads
prove interactive access. Do not rerun unchanged647/UOM228 merely for context.

## 6. Assumptions to re-verify

An asynchronous request asks the user to reconnect RDP without other changes so
native input can be compared. No reply was received before this handoff. The
disposable native diagnostic is `reports/runtime/probe-console-input.ps1`.

## 7. Open questions and blockers

Native cursor control is unavailable in the observed console state; root cause
unproven. New source/observed/interrupted/denied recording tests and all guide
comparisons remain unexecuted. Broader4be-A and final Release1 gates remain open.

## 8. Immediate next action

After native input is verified restored, run `Test-Slice4beGeneralSettings.ps1
-DeployRoot deploy/validation-general-settings-02 -Phase RED -Recorded
-RecordedOnlyDiagnostic`, honoring the error5 stop rule, then resolve actual
failures before the full recorded GREEN retaining all647 ordered baseline checks.

## 9. Critical references

- Architecture v4.11 D18 General Settings catalog31 refinement; Plan022; Controls1.500.
- Handoff120 for accepted647/UOM228 and their exact frozen-candidate evidence.
- `tests/tooling/Slice4beGeneralSettingsRecorded.ps1` and
  `Slice4beGeneralSettingsGuidePaths.ps1`; the controller's new `-Recorded` mode.
- `tests/integration/plan022_slice4be_remaining_acceptance.md`.
