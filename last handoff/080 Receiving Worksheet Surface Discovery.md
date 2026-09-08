# Receiving worksheet surface discovery

## 1. Goal and release outcome

Achieve Release 1 user acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths.
The Goal remains active and incomplete. This session adds 4be.1 worksheet surface
discovery; it does not complete worksheet activity or comprehensive coverage.

## 2. Current verified state

- Last verified 2026-09-08: code main **23297e7**, docs main **83cf95e**, both
  pushed. This handoff/pointer is committed afterward.
- Preserved candidate `deploy/validation-receiving-navigation-identity`: new
  surface-only probe **105/105**, zero duplicate check names, normal exit.
  All five candidate hashes match the prior navigation/identity manifest.
- No runtime VBA, normative contract, package, accepted deployment or NAS change.
  Prior **771/771** and full technical gates remain recorded in handoff 079 and
  its evidence; they were not rerun by this test-only discovery.
- Excel is closed. Code is clean. Preserve unrelated docs: handoff 067 has
  3 additions/3 deletions; critique 023 remains untracked. Neither was staged.

## 3. Decisions and constraints

D18 semantic inheritance remains approved; no new permission is required.
Its existing coverage/exclusion-evidence rule governs this discovery. Plan 022
and controls v1.79 are synchronized without changing Architecture v4.11.

The actual Receiving Ribbon launcher accepts/captures an existing saved operator
workbook whose sole visible sheet is ReceivedTally. The sheet and Confirm Writes
button retain visible settings after save/reopen. Ordinary provisioned/reused
support sheets remain VeryHidden. The worksheet control cannot be excluded on
visibility grounds. Native invocation, explicit activity identity and outcomes
remain pending. No arbitrary macro/service call may impersonate a user click.

All D5/D12/D13/D14, headless authority, exact System_Key, unknown-column,
captured-workbook and packaged launcher reuse invariants remain binding.

## 4. Evidence and traceability

[Surface evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_surface_results.md)
records the exact command, limits and rejected assumptions.

- Final 105/105 protects visibility states, assigned handler, captured workbook,
  unknown header, saved bytes, authority files and unrelated workbook.
- Four initial route failures -> Excel normalized OnAction to unqualified
  `modTS_Received.ConfirmWrites` -> incorrect qualifier assertion, not proven
  routing defect. The corrected assignment test passes.
- Posted-message attempts lacked calibrated handler entry (98/8); foreground
  calibration stopped with a harness exception (92/1). Minimal Excel input
  experiments could not establish foreground ownership; SetCursorPos failed.
  These are not D13 behavioral RED. NativeInvocationVerified remains false.
- PowerShell parsing, whitespace checks, reference links and package hashes pass.
  No new build/compile/layout/static/live-role/chain/UAT acceptance is claimed.

## 5. Do Not Repeat

Do not require a package qualifier in Excel's normalized OnAction string, infer
a route defect from that string alone, or count failed input delivery as RED.
Calibrate a minimal worksheet input fixture before running the longer package
suite. Do not repeat full native experiments while the desktop condition is
unchanged. Never kill unidentified Excel or rebuild loaded packages. Preserve
handoff 079's native dismissal, identity and maintenance constraints.

## 6. Assumptions to re-verify

Git/Excel state and package hashes. Foreground/cursor availability is environment
dependent. Native shape caller/handler entry is unproved; saved visibility and
assigned macro alone do not establish it.

## 7. Open questions and blockers

Worksheet native input remains an evidence limitation, not a full-Goal blocker.
Receiving launcher denial observation, remaining Operations/Admin coverage and
publication can progress independently. Event Tracking Settings, comprehensive
Viewer, recording/conclusions, guide management, both presentations and physical
multi-station/NAS/user UAT remain incomplete. No pending architecture approval.

## 8. Immediate next action

Add a focused packaged test for Receiving launcher capability-denial observation
through the actual generated Ribbon callback, keeping worksheet native proof pending.

## 9. Critical references

- Architecture D18/D13; Plan 022 4be.1-4be.6; controls v1.79; handoff 079.
- `Test-Slice4beConfigCommands.ps1`: all preserved activity switches plus
  `-CheckReceivingSurfaceCoverage -ReceivingSurfaceOnly` for this diagnostic.
- `Slice4beReceivingSurface.ps1`; `Slice4beReceivingLifecycle.ps1` existing actual
  Ribbon/form seams; `tools/build-xlam.ps1` generated capability dispatch.
- `modTS_Received.EnsureReceivingButtons`, `ConfirmWrites`,
  `EnforceReceivingSupportSheetsHidden`, `CanHideWorksheet`.
- Ignored `reports/runtime/slice4be-receiving-activity/surface-reachability.json`,
  `diagnostic-surface-green.json`, `surface-visibility-discovery.log` and earlier
  discovery logs. `surface-native-input-diagnostic.ps1` preserves unsuccessful
  unsaved test instrumentation; it is not a runtime correction or acceptance test.
