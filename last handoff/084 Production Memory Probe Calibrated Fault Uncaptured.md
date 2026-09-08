# Production memory probe calibrated; fault uncaptured

## 1. Goal and release outcome

Achieve full Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths.
Goal remains active. Slice 4be.1 Receiving denial implementation is incomplete as
an acceptance checkpoint because the standard Production regression is unresolved.

## 2. Current verified state

- Last verified 2026-09-08: code main **ebb01aa**, docs main **ea8d952**, pushed;
  this handoff/pointer follows. Controls v1.84. Runtime remains **dec08fe**.
- Three diagnostic full workflows complete 2/2 each. No exception was observed,
  so none supplies the missing fault-region classification or establishes a fix.
- All harness handles terminal; Excel closed; all 20 original/rebuilt/saved/
  compiled package pins unchanged. No runtime, builder, static-baseline, accepted
  deployment, operational workbook or NAS changes.
- Code clean. Preserve unrelated modified handoff 067 (3 additions/3 deletions)
  and untracked critique 023. Neither was edited or staged.
- Prior Receiving 133/133, native 137/137, activity 845/845 and the original
  technical gates retain their scope. See handoffs 081/083 and linked evidence.

## 3. Decisions and constraints

D18 semantic inheritance and critique reconciliation are already approved; no
architecture approval is pending. D5/D12/D13/D14, captured workbook/session,
headless authority, exact System_Key and unknown-column preservation remain.
Calibrated diagnostics are not Product RED/GREEN or a standard-gate exemption.
No speculative native/build/runtime repair was adopted.

## 4. Evidence and traceability

[Memory diagnostic evidence](../../invSys_fork/tests/integration/plan022_slice4be_production_memory_diagnostics.md)
records commands, limits and preserved reports; it links the earlier fault matrix.

- Three absent memory-label checks: 24 PASS/3 FAIL -> observer calibration 27/27.
  Ten live-allocation checks also pass: reserved, committed/private, mapped,
  no-access, guard, free/undefined, invalid-address and bounded output schema.
- QueryMemory uses query-information access and VirtualQueryEx metadata only.
  Result exposes state/type/protection labels, queried/guard flags and a Boolean
  equality of exception/access targets; no raw addresses or page contents persist.
  Existing owned-identity, redaction, exception-preservation and detach checks hold.
- Original candidate native v5: full 2/2; initial observer ends after 180 seconds,
  no exceptions. The full workflow continues outside that window.
- Compiled candidate restart-specific observer: full 2/2; attaches to the newly
  owned Excel process before package loading; no exceptions through restart.
- One bounded original-candidate repeat v6: full 2/2; initial 180-second observer
  again sees no exception. No further identical capture was launched.
- Earlier native v4 first-chance Execute/c0000005 followed by c0000028 remains
  unexplained. Static installed oleaut32 code places return offset 0x9f77f after
  an indirect call at 0x9f77d; supplying VBA procedure remains unknown.
- Four diagnostic parsers, relative evidence links, diffs and statuses checked.
  Package preservation 20/20. No new runtime/static acceptance claim.

## 5. Do Not Repeat

Do not turn instrumented passes into a fixed native defect or repeat the broad
captures unchanged hoping for acceptance. Preserve all earlier failures.
Do not retry exit waits, RibbonX reorder or save/compile preparation as repairs
without new evidence. Workbooks enumeration omits add-ins; the prior boundary
test's exact named lookup was calibrated. Observer v1's swallowed exceptions
remain invalid; retain corrected unhandled-exception behavior.

Before native use, both observer and memory calibration SHA files must match the
helper. Keep owned PID/creation/name checks; no dumps, memory contents, stack
arguments or operational rows. Static DLL code inspection is not process memory.

## 6. Assumptions to re-verify

Git/Excel state, all candidate pins and helper SHA. On 2026-09-08 a read-only check
finds no input-desktop access and no foreground window; native worksheet click
proof remains unavailable. This does not explain the Production fault. Do not
repeat the old foreground/cursor attempt without changed desktop conditions.

## 7. Open questions and blockers

Production's native target and supplying call remain unknown; diagnostics can
still progress, so Goal is neither complete nor blocked. Comprehensive control
coverage, Settings, publication/Viewer, recording/conclusions, guides/comparison,
native worksheet proof and physical NAS/multi-station/human UAT remain open.

## 8. Immediate next action

Calibrate fixed value-free markers around Process Release's quiet-UI, queue,
processor, status-query and refresh boundaries, then combine that finer localization
with the native observer in an isolated unsaved diagnostic of the actual handler.

## 9. Critical references

- `frmProduction.mBtnProcessRelease_Click`, `SubmitProcessAction` (around 6281),
  `modProductionReusableDesigns.SubmitReusableDesignEvent`: existing path.
  Earlier fixed-stage failure stopped at testStage ReleaseSource, after SaveSourceDraft.
- `tools/validate_plan022_packaged_launchers.ps1`: unchanged full workflow.
- Ignored `reports/runtime/slice4be-launcher-denial/`: `NativeFrameProbe.cs`,
  `test-native-frame-probe.ps1`, `test-native-memory-labels.ps1`, both
  `native-*-calibration-helper-sha256.txt` pins and calibration checks;
  `native-stack-probe-v5/`, `native-stack-probe-v6/`,
  `native-restart-memory-probe/`, `memory-diagnostic-package-preservation.json`.
- `create-native-restart-memory-probe.ps1`, generated restart harness,
  `create-stage-probe.ps1` and earlier fixed-stage harness/map: ignored diagnostics.
- Four pinned candidate roots under `deploy/`: `validation-receiving-launcher-denial`,
  `validation-receiving-denial-rebuild`, `validation-receiving-denial-finalized`,
  `validation-receiving-denial-compiled`. Original native v4 evidence remains.
