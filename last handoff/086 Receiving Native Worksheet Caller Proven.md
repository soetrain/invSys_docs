# Receiving native worksheet caller proven

## 1. Goal and release outcome

Achieve full Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths.
Goal and Slice 4be.1 remain active. This checkpoint closes native worksheet entry
discovery, enabling D13 activity coverage through the actual control.

## 2. Current verified state

- Last verified 2026-09-12: code main **49cbe82**, docs main **c09ae0d**, pushed;
  this handoff/pointer follows. Controls v1.86. Runtime remains **dec08fe**.
- Minimal native input **8/8**; packaged surface **115/115**, retaining all prior
  105 check identities, no duplicates. Both visible worksheet cases enter actual
  `modTS_Received.ConfirmWrites` with exact `btnConfirmWrites` caller identity.
- Both worksheet captures inspected. Saved bytes, captured form binding, unknown
  header, authority files and unrelated workbook pass. All 20 package pins unchanged.
- Excel closed; all test sessions terminal. No runtime, architecture, accepted
  deployment, NAS or static-baseline change. Code clean; preserve unrelated modified
  handoff 067 (3 additions/3 deletions) and untracked critique 023, neither staged.
- Prior standard Production/restart **2/2** remains the verified unchanged-package
  GREEN from handoff085, not rerun here. Earlier native crashes remain unexplained.
  Prior Receiving denial 133/133, native 137/137 and full activity845/845 retain scope.

## 3. Decisions and constraints

D18 semantic inheritance and critique reconciliation remain approved. No repeat
architecture approval is pending. This test-only discovery follows existing D18
coverage and D13 rules; it does not implement worksheet activity or invent a new
control ID. Any actual contract change still requires normative approval first.
Keep D5/D12/D14, captured workbook/session, headless authority, exact System_Key,
unknown columns and every accepted regression. No arbitrary service/macro call
may impersonate a user click. Automated captures are not human UAT.

## 4. Evidence and traceability

[Surface evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_surface_results.md)
records exact commands, scope and unsuccessful setup attempts.

- Minimal calibration checks native entry/caller/binding, programmatic distinction,
  wrong-window and inactive-fixture rejection, and two-workbook switching:8/8.
- Packaged native switch adds ten checks to preserved105:115/115. In-situ fixture
  button calibrates delivery before each original button. OnlyStagingVisible and
  SavedReopen report NativeInvocationVerified=True; hidden cases remain false.
- Earlier recovery attempts stop94PASS/1harness failure before real entry. Marked
  captures showed a coordinate miss and then the unrelated sentinel window in front.
  COM parent traversal, active-context and foreground guards also rejected setup.
  None is D13 product RED; no routing repair was justified or implemented.
- Source inspection: worksheet ConfirmWrites calls posting owner directly; form
  Confirm uses modReceivingActivityAction. Worksheet activity needs its own identity,
  context and owner-outcome tests under D18, using this calibrated native path.
- Five PowerShell parsers, changed references, diffs/status and20packagepins pass.
  No new full build/compile/layout/static/live-role/chain acceptance claimed.

## 5. Do Not Repeat

Do not infer input from OnAction assignment or Excel PID ownership alone. Use the
fixture workbook's own window, explicitly show/activate that window, activate its
sheet, and check active workbook/sheet and exact foreground root before clicking.
The passing cases' Window.Visible was already true before preparation; do not
claim a hidden-workbook runtime defect or isolate that assignment as the cause.
Use the calibrated origin/DPI/zoom calculation on an unscrolled fixture.
PowerShell wraps C# exceptions: calibrate rejection using GetBaseException().Message.
Use case-insensitive matching for VBE-normalized identifiers while preserving source.
Do not repeat broad crash diagnostics or speculative runtime fixes without new evidence.

## 6. Assumptions to re-verify

Git state, Excel closure, package pins and available input desktop. The helper is
developer-only and calibrated for the explicit unscrolled fixture; changing input
or window setup requires fresh calibration. Native discovery does not prove any
submitted event, activity record, stale-context guard or Domain effect.

## 7. Open questions and blockers

Worksheet activity identity/refinement and meaningful RED/GREEN remain next.
Remaining Operations/Admin coverage, Settings, publication/Viewer, sequence recording
and conclusions, guides/comparison, physical NAS/multi-station and human acceptance
remain open. Goal is neither complete nor blocked; no external dependency prevents
the next test-first work.

## 8. Immediate next action

Refine the registered worksheet-control activity contract under D18 in the normative
specification, Plan022 and controls, then use the proven native control to establish
missing activity/context/outcome RED before implementation.

## 9. Critical references

- `Test-Slice4beWorksheetInput.ps1`, `Slice4beWorksheetInput.ps1`,
  `Slice4beReceivingNativeSurface.ps1`, `Slice4beReceivingSurface.ps1` under tests/tooling.
- `Test-Slice4beConfigCommands.ps1`: all preserved activity flags plus
  `-CheckReceivingSurfaceCoverage -ReceivingSurfaceOnly -CheckReceivingNativeSurface`.
- `modTS_Received.ConfirmWrites`, `modReceivingActivityAction.ConfirmWrites`,
  `modReceivingPostingService.ExecuteConfirmWrites`; Core `modActivityCatalog` remains6.
- Original candidate `deploy/validation-receiving-launcher-denial`; pins under
  `reports/runtime/slice4be-launcher-denial/` (original/rebuild/finalized/compiled).
- Ignored `reports/runtime/slice4be-receiving-activity/`:
  `native-calibration-d665db225b304b11adfd82b554fe33ab/checks.json` (8/8),
  `native-surface-230a1587ef9d46d2b307ddbce4e11810/` (115/115 plus captures),
  `native-surface-visible-window.log`; earlier native-surface logs retain setup failures.
- Handoff085 and its linked Production recovery record retain the unchanged2/2
  standard gate and historical native-crash limitations.
