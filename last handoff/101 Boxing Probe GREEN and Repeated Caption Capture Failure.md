# Slice 4be: Boxing probe GREEN and repeated caption-capture failure

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11 and Plan 022 without
regressing accepted workflows. Goal remains active/incomplete. This checkpoint
reconciles the stale Boxing test with approved D18; broader capture and clean
shutdown still fail, and multiline rendering remains open.

## 2. Current verified state

Last verified: 2026-09-24, after 19:24 UTC. Both repositories are on main. Code
**78520bc** is committed and pushed. Documentation base is **4117275**; the ensuing
commit contains controls1.243, Plan updates and this handoff. Verify its push.

Frozen candidate remains `deploy/validation-auth-read-separated`; no runtime
source or XLAM changes this checkpoint. Code tree is clean. Preserve unrelated
docs handoff067 edits and untracked critique023. Runtime reports/images are private
and ignored. Excel and all gate controllers are closed; settings are restored.

## 3. Decisions and constraints

D8-A and D18 scrolling were approved on 2026-09-24. Do not reopen approval or
restore Locked=True. ListBox cells remain non-editable while selection/scrolling
is enabled. Boxing's shared probe now checks that contract; native typing/value
preservation remains independently protected. No architecture change is introduced.

The existing submission-only diagnostic now permits Boxing's required recording
probes; it does not claim full recording coverage. Incomplete standalone recording
and owner-completion combinations still reject before Excel starts. Default full
route is unchanged. D12, D13, exact System_Key, unknown columns and captured
binding remain required. Do not edit pinned tooling during a live Excel gate.

## 4. Evidence and traceability

- Focused actual Viewer/Detail: **49 PASS / one expected probe FAIL -> 50/50**,
  retaining all 44 previous checks, adding five instrumented compiles. Four images
  reviewed; native scroll/typing, source bytes and context protected. Both runs
  close normally, restore settings and preserve pins, with zero Application failures.
  This is test-contract reconciliation, not a newly discovered runtime RED.
- Disposable foreground calibration passes hidden-first, visible and restored
  cases; normal closure. It does not explain the long-run capture failure.
- Isolated Shipping/Boxing submission: **794/794**, including all 82 previously
  unreached checks, five compiles and seven reviewed captures. Normal delayed
  closure, restored settings/pins, zero Application events 1000/1001/1002.
- Complete corrected route: **1632 PASS / one capture exception**, 82 checks
  unreached, 1625 prior GREENs retained. Both corrected Make/Unbox ReadOnlyFields
  and all seven D8-A Auth-recreation checks pass. Sixteen scoped images reviewed.
  Same unavailable-store Make capture fails with no uncovered owned caption point;
  failure snapshot shows enabled/non-minimized owned form and VS Code foreground.
  Later native inspection finds no remaining Excel windows during cleanup; this
  does not establish the earlier failure cause. Only disposable Excel is terminated.
  Original host then writes the terminal report and restores settings; host itself
  is not terminated. Normal shutdown is not accepted. Five packages/217 tooling
  pins preserved; zero audited Application failures. UTC19:02:03--19:23:22.
- Static unchanged: 251 components, 6047 procedures, 132972 lines, nine literal/
  45 unresolved dynamic calls, 191 duplicate-body candidates, 28 oversized ratchets;
  three schemas pass and all 277 PowerShell files parse. Source/layout unchanged;
  prior D8-A five builds/compiles and chain/live-role/Create Warehouse32/48/15 retain
  their candidate scope. No package rebuild is needed for this test-only correction.

Private receipts under `reports/runtime/`: `reconciled-boxing-detail-verification.json`,
`boxing-submission-isolation-verification.json`, `reconciled-boxing-detail-static-verification.json`,
and `d8-auth-read-reconciled-regression-boxing-verification.json`.
Full-run report directory: `slice4be-shipping-activity/c527145cf74349daa227962018d7ae5b`.
Its structured `boxing-activity-shipping-recording-green.json` contains a failure;
the filename is not a GREEN claim. Capture/cleanup observations and termination
receipt share the `d8-auth-read-reconciled-` prefix.

## 5. Do Not Repeat

Do not rerun the complete route merely because the isolated capture succeeds.
The same failure now reproduces after the full preceding lifecycle. Observe the
physical hit-test targets, z-order/cloaking and resource state at failure, before
cleanup removes the window. Do not replay business commands to recover a capture.
Do not COM-reattach or issue a second Quit; retain the original settings-restoration
controller. Distinguish assisted closure from normal closure and isolated from
full-route coverage. No elevation requirement or Windows error5 cause is proven.

## 6. Assumptions to re-verify

Desktop access works for calibration, focused and isolated captures, and sixteen
full-route captures. This does not establish sustained capture access. The full
route closes its earlier Viewer/Shipping/workbook surfaces before submission;
that lifecycle/resource state is a hypothesis to investigate, not a confirmed cause.
Check Excel closure and frozen hashes before the next test. One Excel gate at a time.

## 7. Open questions and blockers

Broader caption capture and normal shutdown remain open. Event Detail multiline
readability is separate from horizontal scrolling. A Label in a scrollable frame
was considered as a read-only selected-value presentation; it is tentative, with
no implementation or contract refinement yet. Reconcile it with D18, profiles,
original values and context guards, then establish packaged RED before editing VBA.

Other required work remains comprehensive Operations/Admin tracking (Production
census has 61 pending constructed buttons), guide transfer/provenance, comparison/
presentation, carrier/D5 authority, seven owner image corrections, human/NAS
acceptance. Keep these visible; see the remaining-acceptance checklist and handoff100.

## 8. Immediate next action

Add read-only failure-point capture diagnostics and exercise a focused reproduction
of the preceding window lifecycle before another full Shipping/Boxing retry.

## 9. Critical references

- Architecture v4.11 D18/D8-A; Plan022; controls1.243.
- `tests/integration/plan022_slice4be_detail_overflow_results.md`,
  `plan022_slice4be_auth_read_results.md`, `plan022_slice4be_remaining_acceptance.md`.
- `tests/tooling/Test-Slice4beConfigCommands.ps1`: ActivateByCaptionClick,
  CaptureOwnedFormEvidence, original cleanup and submission-only preflight.
- `tests/tooling/Slice4beShippingActivity.ps1`: full-route cleanup before submission.
- `tests/tooling/Slice4beBoxingTracking.ps1`: unavailable-store Make capture.
- `tests/tooling/Slice4beBoxingPublishedRead.ps1`, `Slice4beViewerEventDetail.ps1`:
  corrected read-only probe and native-input protection.
- `src/Operations/Forms/frmEventDetail.frm`,
  `src/Operations/ClassModules/cEventDetailController.cls`: future multiline work.
