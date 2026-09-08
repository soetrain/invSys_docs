# Production execute fault and package boundary

## 1. Goal and release outcome

Achieve Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths.
Goal remains active. Slice 4be.1 Receiving denial implementation is not a completed
checkpoint: the standard reusable Production/restart gate remains unresolved.

## 2. Current verified state

- Last verified 2026-09-08: code main **8f2e09b**, docs main **d6ee3ac**, pushed.
  This handoff/pointer is committed afterward. Controls v1.83; runtime remains
  **dec08fe**. No runtime, builder or static-baseline changes in this checkpoint.
- Excel closed; diagnostic sessions terminal. All 20 pinned packages across the
  original/rebuilt/saved/compiled sets remain unchanged. No accepted deployment,
  operational workbook or NAS changes.
- Code clean before this handoff. Preserve unrelated docs: modified handoff 067
  (3 additions/3 deletions) and untracked critique 023; neither staged nor edited.
- Original Receiving evidence remains focused 133/133, native 137/137 and activity
  845/845 retaining prior 771 and discovery 105 identities. Earlier compile/layout/
  static/live-role/chain gates retain their original scope; see handoffs 081/082.

## 3. Decisions and constraints

D18 semantic inheritance and critique reconciliation are approved; no architecture
approval is pending. D5/D12/D13/D14, captured workbook/session, headless authority,
exact System_Key and unknown-column preservation remain binding. Diagnostic passes
do not replace the standard full Production gate or human acceptance.

No restart wait or RibbonX reorder was implemented: observed evidence supports
neither repair. The new package-boundary test protects already-correct behavior;
there is no Product RED/GREEN claim for it.

## 4. Evidence and traceability

[Curated exit/execute results](../../invSys_fork/tests/integration/plan022_slice4be_production_exit_and_execute_results.md)
contain commands, limits and exact retained evidence. The linked earlier matrix
preserves other variants and their failures.

- Actual restart teardown: three focused owned-process cases terminal at restart
  boundary. Full compiled-candidate workload with memory-only markers and process
  assertions passes 2/2: terminal first process and distinct restarted identity,
  without a new wait, project inspection or VBA edit. Exit race not observed.
- Native observer metadata calibration: 15 PASS/3 FAIL for absent fields, then
  24/24 including operation decoding, identity, redaction, exception preservation,
  survival and detach. This is tool calibration, not Product RED.
- Original candidate native v4: 0 PASS/1 failure at batch scale. First preceding
  exception c0000005 is first-chance Execute at UNRESOLVED_FRAME, then oleaut32,
  VBE7 and Excel. Later c0000028 retains ntdll/repeated VBE7. Observer detached;
  harness closed Excel. Exact supplying VBA call and memory state remain unknown.
- Live SaveAs calibration: saved XLAM is not loaded; source keeps original name
  and empty path. Exact named-lookup detector calibration 5/5, including explicit
  open rejection and byte preservation. Unchanged real builder passes 7/7 external
  edit boundaries (five identities, two RibbonX). First enumeration/JSON probe
  was inconclusive harness error, not RED.
- Seven scripts parse; evidence links, diffs and Git status reviewed. No new
  claim for remaining runtime/static or visible acceptance gates.

## 5. Do Not Repeat

Do not infer loaded-artifact state solely from SaveAs/Close source ordering.
Workbooks enumeration omits add-ins; use exact named lookup and path comparison.
Do not repeat exit waits, rebuild/save/compile preparation or RibbonX reorder as
repairs without new evidence. Earlier observer v1 could swallow exceptions and
its 2/2 remains invalid. Preserve corrected unhandled-exception behavior.

Native probes require synthetic calibration and pinned helper hash before Excel
use. Verify owned process identity; never persist raw addresses, memory contents,
stack arguments, operational rows or dumps. An unresolved site does not establish
freed memory or identify a VBA procedure. Worksheet input calibration remains
independently pending; do not repeat the failed foreground approach unchanged.

## 6. Assumptions to re-verify

Git/Excel state, candidate hashes and helper SHA. Diagnostic instrumentation can
change execution conditions; the passing exit-proof run does not repair the plain
gate. SaveAs behavior and package detector results describe the calibrated flow.

## 7. Open questions and blockers

Production instability prevents checkpoint acceptance but permits further
diagnostic progress; Goal is neither complete nor blocked. Native worksheet
activity, remaining Operations/Admin coverage, Settings, publication/Viewer,
recording/conclusions, guides/comparison and physical NAS/multi-station/human
acceptance remain open.

## 8. Immediate next action

Calibrate a read-only metadata probe for committed/free state, image/private type
and protection at an invalid execution target, then capture those bounded labels
from an owned native failure without reading or persisting memory contents.

## 9. Critical references

- `tools/validate_plan022_packaged_launchers.ps1`: unchanged full ProductionReusable
  workflow and public launcher/form-action boundaries.
- `tests/tooling/Test-PackagedExternalEdits.ps1`: new observed-copy build test;
  requires fresh validation/evidence directories and Excel closed.
- Ignored `reports/runtime/slice4be-launcher-denial/`: `test-exit-boundary.ps1`,
  `exit-boundary.json`, `create-exit-workload-probe.ps1`,
  `exit-workload-production-probe.ps1`, `compiled/exit-workload/`,
  `NativeFrameProbe.cs`, `test-native-frame-probe.ps1`,
  `native-calibration-checks.json`, `native-calibration-helper-sha256.txt`,
  `native-stack-probe-v4/`, `exit-execute-package-preservation.json`.
- Ignored `reports/runtime/slice4be-closed-package/`: `red/` (invalid probe),
  `red-v2/` (7/7), `saveas-calibration/`, `boundary-guard-calibration/` (5/5),
  `calibrate-saveas.ps1`, `calibrate-boundary-guard.ps1`.
- Pinned candidate roots under `deploy/`: `validation-receiving-launcher-denial`,
  `validation-receiving-denial-rebuild`, `validation-receiving-denial-finalized`,
  `validation-receiving-denial-compiled`. Additional boundary-test package sets
  `validation-ribbon-closed-red` and `validation-ribbon-closed-red-v2` are diagnostic.
