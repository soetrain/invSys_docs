# Slice 4be: caption z-order diagnosis; original multiline values next

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11/Plan022, preserving accepted
workflows and D13 gates. Goal remains active/incomplete. This turn improves capture
input safety and localizes the full-route failure. Proceed with independent D18
original-value/multiline work; retain broader capture and shutdown as open gates.

## 2. Current verified state

Last verified: 2026-09-24, after 20:03 UTC. Both repositories are on main. Code
**273875a** is committed/pushed. Docs base **750c0c2**; the ensuing commit contains
controls1.244, Plan update and this handoff. Verify its push/upstream state.
Code tree is clean after the commit. Preserve unrelated docs handoff067 edits and
untracked critique023. Reports/images remain private and ignored.

Frozen candidate is still `deploy/validation-auth-read-separated`. No runtime
source or XLAM changed. Excel, the diagnostic worker/controller and passive
sampler are closed; original settings restored. Do not restart any former session.

## 3. Decisions and constraints

D8-A and D18 scrolling remain approved. No normative architectural change this
turn. Exact original values, profile restrictions, captured context, headless
authority, exact System_Key and unknown columns remain binding. Future decoder
correction restores the existing contract; it must first have packaged D13 RED.

The capture helper now skips off-monitor points and verifies actual cursor
coordinates before input. Native ownership/hit-test checks and restoration of
original topmost state remain. Failure diagnostics contain only window/process
identities, classes, geometry, cloaking, hit targets and resource counts; never
captions or workbook content. No alternate raise/message flags are implemented.

## 4. Evidence and traceability

- Final disposable calibration **eight cases**: seven reviewed captures (including
  direct caption activation) plus expected offscreen rejection before input. Every
  case restores topmost state. Normal Quit/process closure; zero Application
  1000/1001/1002 events. No invSys package loaded. Earlier offscreen negative
  assertion failed because WindowFromPoint accepted an offscreen owned point and
  the helper did not verify cursor positioning. This is tooling protection, not
  product D13 RED. Final report under runtime:
  `capture-foreground-calibration/9b894feab5814173bdebcb68f9d32d6b`.
- Complete Shipping/Boxing diagnostic **1632 PASS / one capture exception**, all
  preceding reached identities/outcomes retained, 82 full-route checks unreached.
  Five instrumented compiles; both read-only assertions and all seven Auth checks
  pass. Sixteen images produced, not reviewed/promoted to new visual acceptance.
- Failure point: unavailable-store Make. The form is on-screen, visible, enabled,
  non-minimized and not cloaked. SetWindowPos returns success, but AfterRaise still
  has Topmost=False and z-order below VS Code; all three monitor-valid points hit
  VS Code. GDI816 at failure, process peak930; 107 passive samples reach max892
  current GDI and max15 XLMAIN windows. Resource exhaustion/offscreen/cloaking do
  not explain this failure. Why the raise is ignored remains unresolved.
- Terminal report is written before assisted host/Excel termination. Excel is
  briefly enumerable with HasExited=True, then disappears. Passive sampler exits
  without termination. Original controller restores settings; five packages/217
  tooling pins preserved; zero Application failures. Normal shutdown fails.
  Diagnostic UTC19:35:28--19:56:32, report directory under runtime:
  `slice4be-shipping-activity/f16817b7fdff4c4a9dde4ddb061606ee`.
- Static unchanged: 251 components, 6047 procedures, 132972 lines, nine literal/
  45 unresolved calls, 191 duplicate groups, 28 oversized ratchets; three schemas
  pass. All 277 PowerShell files parse. No product/layout rebuild needed for this
  tooling-only change. Earlier candidate chain32/live-role48/CreateWarehouse15,
  focused D8-A82 and Detail50 retain their documented scopes.

Private receipts: `caption-state-calibration-verification.json`,
`caption-state-diagnostic-boxing-verification.json`, `caption-state-static-verification.json`
under `reports/runtime/`. The main evidence record retains every calibration attempt.

## 5. Do Not Repeat

Do not rerun the broad route blindly, equate successful SetWindowPos return with
effective raising, or click a point without ownership/monitor/cursor checks. Do
not infer elevation/error5 or GDI quota failure from this trace. Probe installation
already occurs before forms/five compiles; the later submission installer is guarded.
Do not assume late VBA editing without evidence. Do not COM-reattach or issue a
second Quit during original cleanup; keep its settings snapshot. A phase-named
green.json with a failure is not GREEN.

## 6. Assumptions to re-verify

Capture works in focused/isolated runs but fails after the full lifecycle. Any
alternate z-order strategy must be calibrated on owned windows, restore state,
and retain native visible verification. Its underlying MSForms/Windows cause is
unproven. Confirm Excel closure and frozen hashes before the next packaged gate.
Do not modify pinned tooling during a live gate.

## 7. Open questions and blockers

`frmInventoryViewer.ViewerUnescape` sequentially replaces escaped n/r/t before
escaped backslashes. Core `modPublishedEventsReader.Escape` escapes backslashes
first. A static algorithm control fails preservation for three synthetic literal-
escape/mixed-newline values (`detail-escape-static-control.json`); this is not
packaged RED. Use actual publication, Viewer selection and Event Detail with
literal sequences and real line breaks before changing the decoder. Preserve
source bytes/profile filtering and all previous 50 focused checks. Do not strip
escapes or substitute a hybrid wire contract. Rendering acceptance is separate:
even correct decoded text still needs complete multiline visibility.

Comprehensive Operations/Admin tracking (61 pending Production buttons), guide
transfer/provenance, comparison/presentation, carrier/D5 authority, seven owner
capture corrections and human/NAS acceptance remain. See the maintained checklist.

## 8. Immediate next action

Establish packaged Viewer/Detail RED for literal-escape and real-newline preservation
through the existing Admin-generated published fixture before editing runtime code.

## 9. Critical references

- Architecture v4.11 D18/D8-A; Plan022; controls1.244; remaining-acceptance checklist.
- `tests/integration/plan022_slice4be_capture_lifecycle_results.md`.
- `tests/tooling/Test-Slice4beConfigCommands.ps1`, `Test-Slice4beCaptureForeground.ps1`.
- `tests/tooling/Slice4beViewerEventDetail.ps1`, `Slice4bePublishedProjectionFixture.ps1`.
- `src/Core/Modules/modPublishedEventsReader.bas`: Escape/RenderLine.
- `src/Core/ClassModules/cEventsPublication.cls`: CaptureInventory/FirstToken.
- `src/Operations/Forms/frmInventoryViewer.frm`: ViewerUnescape and actual selection.
- `src/Operations/Forms/frmEventDetail.frm`,
  `src/Operations/ClassModules/cEventDetailController.cls`: read-only fields/context.
