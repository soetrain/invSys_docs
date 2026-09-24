# Slice 4be: D8-A Auth read separation and broader capture blocker

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11 and Plan 022, preserving
accepted workflows. Goal remains active/incomplete. Approved D8-A now has
packaged focused RED/GREEN, five builds/compiles and passing release-chain gates;
broader Shipping/Boxing acceptance remains incomplete.

## 2. Current verified state

Last verified: 2026-09-24, after 18:38 UTC. Both repositories are on main. Code
**88f331d** is committed and pushed. Documentation base is **be5c17a**; the
ensuing documentation commit contains controls1.242, Plan updates and this handoff.
Verify its push/upstream status before ending the turn.

Frozen candidate: `deploy/validation-auth-read-separated`. All five packages
build/compile, Operations cold start passes, and only Core/modAuth changes among
244 compiled components. Accepted deployment and earlier candidates are unchanged.
Code working tree is clean after its checkpoint commit. Preserve unrelated docs
handoff067 edits and untracked critique023. Private reports/images remain ignored.

Excel is closed and all controllers have terminated. Focused/chain gates close
normally; broader Shipping/Boxing requires completed-host and residual-Excel
termination. Its original outer controller restores the in-memory settings snapshot.
Do not describe that broader cleanup as normal shutdown.

## 3. Decisions and constraints

D8-A and Event Detail scrolling were explicitly approved 2026-09-24; no repeated
approval is needed. D8-A ordinary reads must not create, repair or save authority.
Explicit provisioning and D2 rule8's credential-authorized same-user S1 station
transition remain unchanged and are separately protected.

Resolver uses exact selected target, or an explicit Core root with matching
validated Config for existing headless/bootstrap callers. GetCurrentTarget already
suppresses a stale target when its root differs from the explicit transaction root.
No arbitrary open-workbook/default-root discovery remains. Three private helpers
were removed after reviewed caller/dynamic-reference checks. No public signature
changed. D12, exact System_Key, unknown columns and captured binding remain binding.

## 4. Evidence and traceability

- D8-A packaged RED **60 PASS / 22 expected FAIL**, GREEN **82/82**, exact IDs
  retained; five instrumented compiles, normal closure, settings/pins preserved,
  zero Application events 1000/1001/1002. See the Auth evidence record for exact
  controller/report IDs and retained earlier calibration attempts.
- Reconciled Phase 6 test96, `TestLoadAuth_RequiresExistingExplicitlyProvisionedAuthority`,
  passes **1/1**. The superseded auto-bootstrap expectation was removed.
- Current candidate full-chain/live-role/Create Warehouse **32/48/15**, exact
  previous IDs, normal closure, restored settings/three reports, preserved five
  packages/275 tooling hashes and zero Application failure events.
- Static: 251 components, 6047 procedures (-3), 132972 lines (-53), nine literal/
  45 unresolved dynamic calls (unchanged), 191 duplicate-body candidates (-2),
  28 oversized ratchets. Three schemas pass; 277 discovered PowerShell files parse.
  Form/layout source unchanged; preceding source-layout 8/8 and 7/7 retain scope.
- Broader Shipping/Boxing **1630 PASS / three FAIL**, 82 prior checks unreached.
  All seven `Shipping.Access.AuthUnavailable.*.MissingFileNotRecreated` now pass.
  Two failures are Make/Unbox `ReadOnlyFields`: the old probe returns mFields.Locked,
  requiring True despite approved D18 Locked=False. This is a stale test contract;
  do not revert runtime. Third failure is owned-caption capture at
  `boxing.tracking.unavailablestore.make.png`. Sixteen images individually reviewed;
  1623 prior GREENs retained. Partial run is not GREEN.
- Capture diagnostic: owned form visible/enabled/non-minimized, VS Code foreground;
  no uncovered caption sample found. Why raising failed is unproven. Completed
  host and residual disposable Excel were terminated; settings/pins restored,
  zero Application failure events. Clean shutdown remains unaccepted.

Private receipt prefix `reports/runtime/d8-auth-read-`: focused, component,
static and chain `-verification.json`; broader prefix
`d8-auth-read-resume0923-regression-boxing` includes verification and separate host/
Excel cleanup records. Broader report directory:
`slice4be-shipping-activity/1c0e716b3cfb4b0e81bde25c211fab6b` under runtime.

## 5. Do Not Repeat

Do not repeat a broad capture run until its failed precondition is isolated or
changed. Do not replay business commands to recover a capture. Do not COM-reattach
or issue a second Quit during original cleanup; preserve the outer settings snapshot.
The broader script had written its terminal result before assisted host/Excel cleanup.
Do not promote phase-named `green.json` to GREEN when it contains failures.

Use assignment before enumerating ConvertFrom-Json arrays in PowerShell5; wrapping
that pipeline in @() can nest arrays. Serialize review arrays with explicit
ConvertTo-Json -InputObject and sufficient depth. No image pixels were edited.

## 6. Assumptions to re-verify

Desktop capture succeeded for 16 broader images before the caption failure; this
does not establish resolution of earlier Windows error5. Recheck process closure,
desktop state and frozen package hashes before further Excel work. One Excel gate
at a time; do not edit pinned tooling or rebuild packages during a live gate.

## 7. Open questions and blockers

Reconcile Boxing read-only probe; isolate caption activation and cleanup failure;
finish broader candidate acceptance. Other Slice4be requirements remain: multiline,
comprehensive Operations/Admin tracking (Production census still 61 pending
constructed buttons), guide transfer/provenance, comparison/presentation,
carrier/D5 authority, seven prior owner image corrections, human/NAS acceptance.
Previous Detail/Viewer/owner/Production results retain their recorded candidate
scope; see handoff099 and the maintained remaining-acceptance checklist.

## 8. Immediate next action

Reconcile the Boxing read-only probe with approved D18 and establish a focused
packaged native-input/caption calibration before another broad Shipping/Boxing run.

## 9. Critical references

- Architecture v4.11 D8-A, D2 operator sign-in rule8 and D18 scrolling decision.
- Plan022, controls1.242, `plan022_slice4be_remaining_acceptance.md`.
- `tests/integration/plan022_slice4be_auth_read_results.md`.
- `src/Core/Modules/modAuth.bas`: LoadAuth, ResolveAuthWorkbook,
  CloseTransientAuthAfterLoad, existing explicit provisioning/transition procedures.
- `tests/tooling/Slice4beAuthReadOnly.ps1`, `Test-Slice4beAuthReadOnly.ps1`.
- `tests/tooling/Slice4beBoxingPublishedRead.ps1`: BoxingDetailForTest ReadOnly
  branch and Make/Unbox ReadOnlyFields assertions; not yet corrected.
- `tests/tooling/Test-Slice4beConfigCommands.ps1`: ActivateByCaptionClick,
  CaptureOwnedFormEvidence and original cleanup; capture failure JSON in report.
