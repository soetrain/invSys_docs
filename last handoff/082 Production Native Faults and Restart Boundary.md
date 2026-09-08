# Production native faults and restart boundary

## 1. Goal and release outcome

Achieve full Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths.
Goal remains active. Slice 4be.1 Receiving denial implementation is not a completed
checkpoint: the standard reusable Production/restart gate remains unresolved.

## 2. Current verified state

- Last verified 2026-09-08: code main **6f0c2c6**, docs main **314b5e2**, pushed.
  This handoff/pointer is committed afterward. Controls v1.82; Architecture D18
  remains approved and unchanged. Runtime source remains **dec08fe**.
- Excel closed; every diagnostic session terminal. All 20 package hashes across
  original, clean rebuild, saved-copy and compile-before-save candidates match
  their respective pre-test manifests. No deployment/NAS/operational changes.
- Code clean. Preserve unrelated docs: handoff 067 remains 3 additions/3 deletions;
  critique 023 remains untracked. Neither was staged or edited.
- Original Receiving denial evidence remains focused 133/133, native 137/137,
  activity 845/845 retaining all prior 771 and discovery 105 identities. Prior
  build/compile/cold-start, smoke 86, live 48, chain 30, Viewer/layout and launcher 3/3
  gates remain scoped to their original candidate; see handoff 081/evidence.

## 3. Decisions and constraints

D18 semantic inheritance and critique reconciliation are already approved.
No new architecture approval is pending. D5/D12/D13/D14, captured workbook/session,
headless authority, exact System_Key and unknown-column preservation remain.
No speculative runtime/build repair or regression exemption was adopted.
Diagnostic success never substitutes for the unchanged full gate or human UAT.

Read-only provenance probes do access VBProject; normal runtime must not gain
that requirement. Native and source-stage probes alter execution conditions and
remain diagnostics. Keep all variants and failures distinct.

## 4. Evidence and traceability

[Curated matrix](../../invSys_fork/tests/integration/plan022_slice4be_production_provenance_diagnostics.md)
contains exact variants, package hashes, calibration, limits and retained paths.

- Provenance variants: combined candidate Core+Operations/prior Domains, relocated
  all-current executing packages, and original files/location each pass 2/2 with
  verified initial/restart dependency paths. Admin is not loaded by this harness.
- One-second post-load pause without project access:0/1 at batch scale.
  Resolved package save-only copy:0/1 there despite source 168 unchanged and all
  five compiles/cold-start PASS. Neither supports a build fix.
- Corrected native observer v2: target c0000028 captured, 0/1 at variable quantity.
  v3 captures preceding observed c0000005 with unresolved first frame, then
  oleaut32/VBE7/Excel; target stack includes ntdll/repeated VBE7. Same boundary 0/1.
  Exact VBA trigger remains unknown. Both observers detach without engine error.
- Native calibration extended from 13 PASS / 2 FAIL (missing preceding code/frames)
  to 15/15; identity mismatch rejection, redaction, non-target exception preservation,
  fixture survival and detach all pass. This is tool calibration, not Product RED.
- Unsaved fixed-stage probe: 0/1; last marker ReleaseSource in released Process
  edit/export test, after SaveSourceDraft and before ViewSource. No saved VBA edits.
- Compile-before-save copy: force all five project compiles then save, source 168
  unchanged; plain full harness **1 PASS/1 failure** during clean restart.
  Same set plus 20 fixed PowerShell restart markers: **2/2**, no VBProject access
  or VBA edits. Markers reach post-restart-action inspection. Cause unresolved.
- Original static baseline unchanged; no new runtime/static GREEN claimed.
  Eight diagnostic PowerShell parsers, links, diffs and statuses checked.

## 5. Do Not Repeat

Do not accept the first native observer's 2/2: it could swallow other exceptions.
Its corrected versions resume non-target exceptions as unhandled. Native helpers
must pass synthetic calibration and match the pinned helper SHA before Excel use.
Use only owned PID+creation identity; no dumps/raw memory/arguments/row values.
Do not repeat plain rebuild, Operations-only recompile, save-only preparation or
all-project compile/save as a claimed repair. Do not add runtime VBProject access.
Do not infer a Domain cause from substitution success; the all-current control
also passed. No changes to standard validator or build script were committed.
Handoff 080's uncalibrated worksheet-input attempts remain independently pending.

## 6. Assumptions to re-verify

Git/Excel state and candidate hashes. The standard restart block requests Quit,
releases automation, waits 750 ms and stops a remaining first process; it does not
explicitly WaitForExit before constructing the next Excel instance. Whether this
permits an exit-order race is **unverified**, not an established root cause.
The restart marker success does not localize the earlier plain failure.

## 7. Open questions and blockers

Production stability blocks this checkpoint's acceptance, not further diagnostic
progress. Native worksheet activity, remaining Operations/Admin coverage, Settings,
publication/Viewer, recording/conclusions, guides/comparison and physical NAS/
multi-station/human acceptance remain open. Goal is neither complete nor blocked.

## 8. Immediate next action

Create a focused identity/terminal-state diagnostic proving whether the outgoing
owned Excel process has exited before restart construction, without VBProject
inspection, then use that evidence before any harness correction or full rerun.

## 9. Critical references

- `tools/validate_plan022_packaged_launchers.ps1`: full ProductionReusable path,
  restart block around 1360-1502; public Production launcher and restart action.
- `reports/runtime/slice4be-launcher-denial/` (ignored): provenance variants,
  `finalized/`, `native-stack-probe-v2/`, `native-stack-probe-v3/`,
  `fixed-stage-probe/`, `compiled/production/`, `compiled/restart-stages/`,
  `final-diagnostic-package-preservation.json`.
- Ignored `NativeFrameProbe.cs`, `test-native-frame-probe.ps1`, calibrated SHA;
  `create-stage-probe.ps1`, `create-compiled-candidate-probe.ps1`,
  `create-restart-stage-probe.ps1` and their generated exact harness variants.
- Candidate roots: `deploy/validation-receiving-launcher-denial`,
  `validation-receiving-denial-rebuild`, `validation-receiving-denial-finalized`,
  `validation-receiving-denial-compiled` (all under deploy, all ignored).
- Prior accepted `deploy/validation-receiving-navigation-identity`; handoffs 080/081.
