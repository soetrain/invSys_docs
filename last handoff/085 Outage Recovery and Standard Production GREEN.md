# Outage recovery and standard Production GREEN

## 1. Goal and release outcome

Achieve full Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths.
Goal remains active. Slice 4be.1 coverage and visible/human acceptance remain
incomplete; the unchanged standard Production gate now passes after recovery.

## 2. Current verified state

- Last verified 2026-09-12: code main **f323bc6**, docs main **83fb755**, pushed;
  this handoff/pointer follows. Controls v1.85. Runtime remains **dec08fe**.
- Unmodified full ProductionReusable validator against the original candidate:
  **2/2**, including clean restart. No debugger, markers, project inspection or
  reduced-workflow flags. Earlier native crashes remain unexplained; no fix claimed.
- Separate corrected Release-marker diagnostic: full 2/2. All resumed handles
  terminal; Excel closed. All 20 pinned original/rebuilt/saved/compiled packages
  unchanged. No runtime, builder, static baseline, accepted deployment, operational
  workbook or NAS changes.
- Code clean. Preserve unrelated modified handoff 067 (3 additions/3 deletions)
  and untracked critique 023; neither edited nor staged.
- Prior Receiving denial focused 133/133, native 137/137, activity 845/845 and
  original technical gates retain their scope; see handoffs 081/084 and evidence.

## 3. Decisions and constraints

D18 semantic inheritance and critique reconciliation are approved; no architecture
approval is pending. D5/D12/D13/D14, captured workbook/session, headless authority,
exact System_Key and unknown-column preservation remain binding. The current
standard-gate pass is real acceptance-test evidence, but it does not prove the
cause of prior native failures, a runtime repair, or human acceptance.

The outage/recovered desktop is an observed environment change, not an established
cause. Keep historical failures visible when evaluating future regressions.

## 4. Evidence and traceability

[Release/recovery evidence](../../invSys_fork/tests/integration/plan022_slice4be_production_release_boundary_diagnostics.md)
contains commands, results and links to earlier native/exit/memory investigations.

- Before outage: placement 2 PASS/11 FAIL ->13/13; disposable marker transport5/5,
  including exact ordering, bounds and preservation of VBA error state.
- Original diagnostic actually stopped during setup on 2026-09-08, 0 PASS/1
  harness failure, before callbacks. On recovery its handle was absent and Excel
  closed. Do not infer work still running from chat/progress files.
- Symptom: boundary101 not found. Read-only loaded VBA showed `.Name` normalized
  to `.name`. Focused loaded-source0/3 ->3/3 after case-insensitive matching while
  preserving exact source; placement13/13 and fresh transport5/5 still pass.
  This is a diagnostic correction, not Product RED/GREEN or a header contract change.
- Corrected full diagnostic2/2: exactly204 numeric markers, comprising11 complete
  Release sequences plus two three-marker calibrations. Native observer has no
  exception during its bounded initial window and detaches without error.
- Native27/27 and memory10/10 calibrations rerun after recovery. Standard full
  gate then passes2/2 unmodified. Six diagnostic parsers, links/diffs/status checked;
  final package preservation20/20. No new runtime/static acceptance claim.

## 5. Do Not Repeat

Do not use case-sensitive matching for VBA identifiers normalized by Excel, change
their saved source merely to satisfy a diagnostic, or count setup failures as RED.
Do not repeat broad native captures hoping for acceptance; the standard gate has
now been run directly in the recovered environment. Retain earlier failures.
No speculative exit-wait, RibbonX reorder, save/compile or mouse-hook repair was
established. Native observers still require both calibrated helper pins and owned
PID/creation/name checks; no raw addresses, page contents, arguments or dumps.

## 6. Assumptions to re-verify

Git/Excel state, candidate pins and helper SHA. Input-desktop access and foreground
window are now available (2026-09-12); both were absent on2026-09-08. That permits
fresh native-input calibration but does not prove worksheet-button handler entry.
Do not jump straight into the long native suite before the minimal fixture works.

## 7. Open questions and blockers

Native worksheet caller proof, remaining comprehensive control coverage, Settings,
publication/Viewer, recordings/conclusions, guides/comparison and physical NAS/
multi-station/human acceptance remain open. The earlier native cause is unresolved.
Goal is neither complete nor blocked; no additional approval is pending.

## 8. Immediate next action

Calibrate actual worksheet-button input in a minimal owned disposable workbook on
the restored desktop, then run the packaged Receiving native-caller test through
the real worksheet control before implementing any missing observation behavior.

## 9. Critical references

- `tests/tooling/Test-Slice4beConfigCommands.ps1`, `Slice4beReceivingSurface.ps1`;
  ignored `reports/runtime/slice4be-receiving-activity/surface-native-input-diagnostic.ps1`.
  Handoff080 records the former delivery/calibration failures, not Product RED.
- `tools/validate_plan022_packaged_launchers.ps1`: unchanged full standard gate.
- Ignored `reports/runtime/slice4be-launcher-denial/`: `post-outage-standard-production/`,
  `post-outage-package-preservation.json`, `release-boundary-probe/` (setup failure),
  `release-boundary-probe-casefix/`, `release-boundary-loaded-procedure.txt`,
  `test-release-marker-live-case.ps1`, `test-release-marker-placement.ps1`,
  `test-release-marker-transport.ps1`, `release-marker-transform.ps1`,
  `ReleaseBoundaryProbe.bas.txt`, `create-release-boundary-probe.ps1` and generated harness.
- Four pinned roots under `deploy/`: `validation-receiving-launcher-denial`,
  `validation-receiving-denial-rebuild`, `validation-receiving-denial-finalized`,
  `validation-receiving-denial-compiled`. Prior native v4 execute-fault evidence remains.
