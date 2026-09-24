# Slice 4be: Production draft tracking checkpoint

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11 and Plan022, with
comprehensive Operations/Admin events and both Action Path presentations.
**Goal active; Slice 4be and six-control acceptance incomplete.** This checkpoint
implements the first six Production draft observations under D18.

## 2. Current verified state

Last verified: 2026-09-24 04:13:54 UTC. Both repositories main; code **f8ea273**,
docs **190907a** before this handoff; both pushed. Controls **1.237**. Current
candidate `deploy/validation-production-designer-captions`; accepted deployment
unchanged. Excel closed; all workers terminal. Five candidate hashes preserved.

Catalog 12 has 68 IDs, preserving versions 1-11. Core adds modProductionDraftCodes;
frmProduction routes six actual New/Clear/Validate handlers through a shared
captured-context/capability guard. Local STAGED/VALIDATED/REJECTED never asserts
saved or Domain-applied designs. Captions match existing buttons exactly.

Preserve unrelated edits: code frmEventDetail.frm +18 measurement lines, still
Locked=True; docs handoff067 +3/-3 and untracked critique023. The existing Event
Detail bytes remain in candidate binaries but were not committed. No other task
edits remain uncommitted after this handoff is saved.

## 3. Decisions and constraints

The six controls inherit D18; the normative specification, Plan022 and controls
were updated before runtime implementation. No authority/permission amendment
or hybrid is introduced. +45 lines in oversized frmProduction have the explicit
scoped Plan022 maintenance exception; total runtime growth is 124 lines/+5
procedures/+1 component. Dynamic calls and duplicate-body counts do not grow.

Pending, not effective: D8-A Auth read/provisioning separation and Event Detail
Locked=False for non-editable selection/scrolling. D5 carrier authority conflict
and exact guide-transfer wire/provenance contract remain unresolved. Do not infer
approval from goal continuation. User permits closing Excel; preserve unsaved work.

Current capture succeeds without elevation: test shell not Admin at 04:03:06 UTC,
all ten Settings images readable/unobscured. Earlier error5 cause remains unknown;
console/RDP switching is a hypothesis only. No permissions/security change made.

## 4. Evidence and traceability

Exact receipts/calibration limitations are in code repository
`tests/integration/plan022_slice4be_production_designer_results.md`.

- Initial actual-handler RED 63 PASS/165 expected FAIL -> GREEN 228/228.
- Caption refinement RED 231 PASS/3 expected FAIL -> GREEN 234/234; all prior
  identities retained. Settings/packages preserved; eventual unassisted closure.
- Five-package build/explicit compile/cold start pass. 244 compiled components;
  initial feature changes modActivityCatalog/frmProduction and adds the vocabulary;
  caption correction changes only that vocabulary module.
- Static: 251 components/6050 procedures/133021 lines; 9 literal/45 unresolved
  Application.Run, 193 duplicate bodies. Three schemas valid; 270 PowerShell files
  parse. Existing Production layout source gates 7/7 and 8/8 retained.
- Current Settings 202/202, exact prior IDs, ten reviewed images, immediate
  unassisted closure, 211 test/five package hashes and settings preserved.
- Current reusable Production aggregate PASS, exact 67 Boolean values/order
  retained, settings restored. Existing helper can terminate after Quit, so
  strict unassisted closure and native Production visibility are not proved.
- Current full chain/live roles/Create Warehouse 32/48/15, exact prior IDs,
  unassisted closure, 269 tooling/five package hashes preserved, settings/three
  reports restored, zero Application 1000/1001/1002 events at 04:13:54 UTC.

Receipts under reports/runtime use `production-designer-captions-` prefixes:
focused-verification, compiled-settings-verification, reusable-verification,
chain-verification, component-comparison and static-verification JSON.

## 5. Do Not Repeat

Do not rerun accepted 234/202/reusable/chain gates without a relevant change.
Preserve calibration failures as harness failures; do not call them full RED.
Older 191 Settings wrapper omitted 11 compile/install checks; use the compiled 202
wrapper. Historical 780 Settings diagnostics were not rerun on this candidate.
Do not infer error5 resolution or elevation requirements from a transient session.
Do not attach another COM client or reissue Quit after the chain's original Quit.
Do not fabricate Recipe nodes or owner completion from matching projection status.

## 6. Assumptions to re-verify

Check Excel/desktop state and candidate hashes before dependent work. Sensitive
settings snapshots stay in memory until processes exit. Run one Excel gate at a
time; no builds while relevant books/add-ins are open. Private runtime reports
and captures remain ignored; no sensitive fixtures or generated reports committed.

## 7. Open questions and blockers

Positive Recipe Validate needs a real released-Process fixture and actual-handler
VALIDATED observation proof. Recording/published Viewer/diagnostic integration
and native Production images remain open for these six controls. Save/Release/
Obsolete still require exact queued-event correlation, including rejection with
pre-existing matching projection status. Other reachable Operations/Admin/shared
controls remain unaccounted for; use the remaining-acceptance checklist.

Guide Create fixture dispatch, comparison Applied label 0x800AC472, guide transfer,
pending amendments, current Viewer/guide acceptance, human comparison and NAS
acceptance remain open. Earlier crashes are unexplained; current chain success
does not close those historical investigations or full Slice 4be acceptance.

## 8. Immediate next action

Extend the packaged Production designer test with a genuinely released-Process
fixture and verify the actual Recipe Validate handler's VALIDATED observation,
establishing focused RED on the pre-observation packages before any needed fix.

## 9. Critical references

- Code remaining checklist and Production designer evidence named above.
- Test-Slice4beProductionDesigner.ps1, Slice4beProductionDesignerActivity.ps1 and
  Slice4beProductionDesignerProbe.ps1 under tests/tooling.
- frmProduction ValidateRecipeDraft(report,True), RecipeRequirementsResolved,
  ProcessRecordsForRecipeNode, mBtnProcessSave_Click/mBtnProcessRelease_Click,
  mBtnRecipeAddProcess_Click, mReusableActionTestInProgress; modOperationsPrimitiveBridge
  and AddSelectedReleasedProcessToRecipe. Use fresh generated identity and actual
  saved/released authority, with setup complete before preservation snapshots.
  Set DesignsEnabled explicitly through the authorized fixture setup command;
  do not use a legacy fallback or replace the owning validator.
- Handoff094 for exact guide Create fixture failure; historical acceptance stays
  attached to its recorded candidate and scope.
