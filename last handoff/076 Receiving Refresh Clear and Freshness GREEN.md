# Receiving Refresh, Clear and freshness candidate GREEN

## 1. Goal and release outcome

Achieve Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths,
without regressing accepted workflows. The full Goal remains active.

## 2. Current verified state

- Last verified 2026-09-07: code main **130df29**, pushed; docs main **896755d**,
  pushed. This handoff/pointer is committed afterward.
- Active **4be.1 Coverage and activity foundation**, incomplete. The Receiving
  Refresh/Clear/freshness checkpoint passes **491/491** focused checks and all
  checkpoint release gates. Controls v1.70 records candidate status.
- Candidate: `deploy/validation-receiving-freshness`; exact five hashes are in
  the linked evidence. Accepted `deploy/current` and NAS are not rolled out.
  Excel is closed; recheck before builds/validation.
- Preserve unrelated docs changes: handoff 067 has 3 additions/3 deletions;
  critique 023 remains untracked. Neither was edited or staged.
- Pre-freshness code **3b06425** and `deploy/validation-receiving-local` retain
  reproducible freshness RED. Earlier catalog-3 candidate remains in
  `deploy/validation-activity`; comparison worktree `../invSys_staging_baseline`
  remains at **f813f2f**. Do not relocate built XLAMs with absolute references.

## 3. Decisions and constraints

Approved D18 synthesis and semantic inheritance remain binding; no approval
blocker. The precise freshness refinement was synchronized in Architecture,
Plan and controls and pushed as docs **bd15052** before implementation.
Discovered controls/outcomes inherit D18; contradiction or material weakening
still requires an explicit approved architecture decision.

Catalog 4 registers eight controls: two configuration actions, four Receiving
Add/Confirm actions, and shared Refresh/Clear. Catalogs 1-3 remain readable;
older policies cannot implicitly enable new controls. Comprehensive coverage
is not implemented.

Core's read-model owner returns optional primitive REFRESHED/STALE/FAILED state
through the existing bridge. True retains its compatible meaning and may include
cached/stale data. STALE is Warning/Changed for local projection/metadata only;
the actual form preserves the owner's cause. Clear distinguishes CLEARED,
EMPTY and uncertain/partial FAILED. Both controls have empty business references.
Optional observation cannot retry/block/redirect authorized work. Captured
session/workbook guards precede the owner independently of optional tracking.
Preserve exact System_Key, unknown columns, D5 ownership, headless authority and
D12's five packages.

## 4. Evidence and traceability

See [checkpoint evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_local_results.md)
for exact commands, hashes, results and limitations.

- Refresh/Clear RED: 365 PASS / 74 FAIL, retaining earlier 262 GREEN; initial
  implementation 439/439 GREEN. Real cached/fallback source cases then produce
  463 PASS / 28 FAIL against that candidate, followed by **491/491 GREEN**.
  All independent identity, unknown-value, authority-byte and binding guards pass.
- Five builds/explicit compiles/cold start; corrected packaged **86/86**;
  live-role **48/48**; ordered full chain/restart **30/30**; Viewer; three-size/
  five-page Production layout/window transitions; launchers **3/3**; dedicated
  reusable Production/full Chai cases/clean restart **2/2**, without reduced flags.
- All 22 local/freshness form captures and three Production layout captures were
  inspected. Automated visible evidence does not constitute human acceptance.
- Final five candidate and five pre-fix hashes unchanged. No Excel process
  remains; final layout/launcher interval has no Excel Application Error event.
- Static: 171 components, 5,465 procedures, 1,077 candidates, 192 duplicate groups,
  45 unresolved dynamic calls, eight literal targets. All 28 current oversized
  limits hold; read-model owner shrinks 1,919 -> 1,916 lines.

The old packaged validator reported 81 passing checks but left a Designs save
prompt: it lacked runtime-root isolation and omitted smoke-opened workbooks from
cleanup. The identified test-session save was declined and Excel exited. The
saved file timestamp predates the run; no initial byte baseline existed. Corrected
validation binds the temporary root in both sessions, provisions Config through
the existing Core setup boundary, refuses out-of-scope workbook cleanup and
requires a real empty collection before ending its HWND-identified process.
Intermediate 84/85 exposed the old external Config dependency; final 86/86
retains all 81 prior checks plus five guards. Saved default Designs bytes are
unchanged across the corrected runs. These are harness findings, not product RED.

## 5. Do Not Repeat

- Do not treat null Excel COM properties as an empty workbook collection or
  terminate an unidentified/user process. Run Excel validators serially.
- Keep the packaged validator's temporary Config/root and complete cleanup
  guards; do not restore ambient warehouse dependencies or tracked-list-only
  shutdown. Do not save an unexpected operational workbook to clear a dialog.
- Preserve cooperative Production dialog-observer shutdown. Earlier native
  ntdll/RPC failures occurred on both older/candidate runtimes; cause remains
  unproven. This candidate's full reusable/restart run passes.
- Historical `Test-Slice12ReviewedCleanup.ps1` has six pre-existing failures on
  both candidate and f813f2f; it is not current GREEN evidence. Pass explicit
  RepoRoot to avoid its parameter-binding failure. Current static limits were
  independently compared with HEAD and hold.
- Do not infer freshness from Boolean True or parse owner error/report text.
  `DiagnoseInventoryReadModelRefresh` actually refreshes and is not a read-only
  diagnostic probe. Keep source values and runtime reports/captures ignored.

## 6. Assumptions to re-verify

Git/Excel state, candidate hashes, current callbacks and fixture isolation.
Physical multi-station/NAS UAT and human comparison remain separate requirements.

## 7. Open questions and blockers

No current approval/environment blocker. Remaining Receiving Open/Close,
optional navigation/selection and hidden worksheet-button reachability need
coverage; other Operations/Admin surfaces need comprehensive mapping and tests.
Publication, Event Tracking policy/detail Settings, personal preferences,
comprehensive Viewer, recording/conclusions, guide management/export/import and
Compare both remain incomplete. Required canonical audits remain effective.

## 8. Immediate next action

Add and run packaged actual-callback/handler RED for the remaining Receiving
Open/Close control coverage while retaining the 491 GREEN baseline, before
extending catalog definitions or changing launcher/form behavior.

## 9. Critical references

- Architecture v4.11 D18 Refresh/Clear/freshness and semantic inheritance;
  D5/D12/D13/D14; Plan 022 4be.1-4be.6; controls v1.70 coverage map.
- `Test-Slice4beConfigCommands.ps1` with activity evidence/foundation, Receiving,
  staging and local switches; `Slice4beReceivingLocal.ps1`,
  `Slice4beReceivingFreshness.ps1`; exact command is in checkpoint evidence.
- `frmReceiving.RefreshClicked` / `mBtnClear_Click`,
  `modReceivingActivityAction.LocalAction`, `modReceivingActivityCodes.LocalOutcome`,
  `modTS_Received.RefreshReceivingUiForWorkbook`,
  `modOperationsPrimitiveBridge.RefreshInventoryReadModel`,
  `modOperatorReadModel.RefreshInventoryReadModelForWorkbook`.
- `tools/validate_phase6_packaged_xlams.ps1`: `Set-PackagedValidationRoot`,
  `Close-PackagedValidationSession`; `tools/validate_plan022_packaged_launchers.ps1`.
