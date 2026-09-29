# Desktop stop and Process component RED continuation

## Goal and release outcome

Complete Release1 acceptance under Architecture v4.11 / Plan022, including Slice4be
comprehensive Event Viewer and Action Paths. **Incomplete; goal paused2026-09-29**
under the user's explicit instruction to stop on desktop error5. Do not resume
tests or implementation without user resume.

## Current verified state

Last verified:2026-09-29. Code `main` **fc359c6**, documentation `main` **9c94f3d**
are the pre-handoff heads; this handoff/pointer is the following documentation
commit. Code is clean. Preserve unrelated modified
`last handoff/067 Partial Goal Action Path NAS Contract.md` and untracked
`expert guidance docs/023 Slice 4be Critique.md`; both still match their private pins.
No Excel process remains. The owned finite desktop observer is stopped.

Frozen candidate: **deploy/validation-production-uom-extent**, runtime source
**dbc74d5**, catalog15. It is not promoted. Current controls1.272, coverage1.15.
Architecture specifies the next catalog16 requirement/output group under D18
semantic inheritance (documentation4dbbf76); no component runtime implementation
has started. Its packaged RED baseline remains unestablished due to fixture errors.
Do not rebuild/deploy while relevant Excel files are open; reverify the five pins
in `reports/runtime/production-component-controller/5aafcb64716b44a89a65450d686babd7/package-pins.json`.

## Decisions and constraints

- Approved UOM reuse is implemented: existing drafts/custom columns survive;
  saved catalog loads only for a new empty workbench. No implicit refresh/save.
  Retrieve/unlist/reopen preserves the complete owned extent, including internal
  blank rows, via hidden worksheet-local `_invSysUomDraftExtent` metadata.
- D18 permits discovered controls within approved rules without repeat approval;
  a contradiction or material weakening still needs an explicit approved decision
  in Architecture first. Catalog16 preserves existing Update identity/fallback/
  append and validation-time editor changes, no-selection Remove reset, quantity/
  UOM and regulation/move semantics. Only STAGED concludes CommandCompleted;
  saved authority Unchanged and empty references never imply Domain application.
- Preserve captured-workbook/session binding, immutable exact `System_Key`, unknown
  columns, headless Core/Domain authority, packaged launcher reuse and every prior
  GREEN identity. Runtime forms/handlers are not replaced by test-only success paths.
- User permits closing Excel; no reboot is authorized. Keep invSys.StationUpdate
  disabled. Do not change lock policies or infer that full access unlocks Windows.

## Evidence and traceability

Current extent candidate; original candidate scopes remain explicit in evidence:

| Gate | Verified result / qualification |
| --- | --- |
| UOM observation | Earlier observation candidate RED98/134 ->232/232; extent preservation RED96/14 -> visible110/110. |
| Combined visible UOM | **264/264**, all prior identities, five compiles, normal unassisted closure, settings/packages preserved, two principal images reviewed. Only the retained-form fixture activates a separate decoy at construction; original refusal assertions remain. |
| Actual public UOM launcher | **61/61**,19 focused plus42 shared, five compiles, normal closure and three principal images. Owner close disposes real form; public reopen retains saved draft/custom column and exact REUSED pair. |
| UOM / Instruction paths | **84/84 /105/105**, exact identities, six principal images per route, normal closure; UOM Detail selects exact terminal REUSED RecordId. |
| Current regressions | Settings202/202, Instructions411/411, draft/paths390/390, lifecycle615/615, native cancellation94/94; exact prior identities, normal closure and preservation. Four native dialog captures reviewed. |
| Settings observations | **468/468**, all467 prior identities plus catalog15 UOM exclusion; five compiles/eight images, preservation. Excel exits unassisted about118 seconds after report. Do not infer a delayed-exit root cause or erase earlier assisted runs. |
| Build/layout/static | Five builds/compiles/cold load;249 compiled components, only Operations UOM catalog source changes from preceding observation candidate. Layout three sizes/five pages/native actions and complete images pass. Static256 components/6073 procedures/133570 lines,9 literal/45 unresolved calls,191 duplicate groups,28 existing caps retained. Three schemas pass. Latest309 PowerShell files parse. |
| Smoke |86/86 with exact identities/preservation; normal shutdown unobserved because validator can force termination without recording its branch. |
| Full chain/reuse | Current extent candidate has no accepted full-chain/reuse proof. Earlier UOM-staging candidate fails Shipping Sent/nativec0000028; earlier cold reusable candidate fails batch scale. Do not carry an older32/48/15 chain/live/Create result forward. |

Authoritative detailed UOM/current-regression evidence:
`tests/integration/plan022_slice4be_production_uom_activity_results.md` in code.
Latest verification receipts (ignored `reports/runtime/`):
`production-uom-extent-visible-host-verification.json`,
`production-uom-extent-public-close-verification.json`, and
`production-uom-extent-settings-activity-verification.json`.

**Pending component fixture:** the first attempt compiles five projects and reaches
48 PASS/90 FAIL, including a harness failure: test snapshot reads unused output
slots and encounters VBA80070057. Owned Excel termination restores settings/packages.
Snapshot now reads the seven requirement/ten output fields and returns explicit
probe errors, but that correction is unverified. The next attempt stops1 PASS/1
harness failure at partial-write probe anchor setup before compilation/handlers;
it closes normally and preserves settings/packages. Neither is protecting RED.
See `tests/integration/plan022_slice4be_production_component_results.md` for exact
controllers/results. Do not start runtime work until a complete meaningful RED.

**Desktop stop:** first recorded error5 after the12:52 Pacific baseline occurs
**15:34:46.392 Pacific /22:34:46.392 UTC**; last successful sample15:34:43.368.
CursorError=5, InputDesktopError=0, CaptureError=6. This is2h42m46s after the user's
baseline, not a measured idle timeout. A logger-sharing failure creates an explicit
gap15:11:09--15:13:12; shared-access logging resumes afterward. The second component
controller already closed Excel at22:34:40 before the desktop failure.
Private receipt: `reports/runtime/desktop-lockout-onset-20260929-153446.json`;
cumulative record: `desktop-lockout-timing-20260929.json`. The observer is stopped.

## Do Not Repeat

- Do not rerun accepted broad gates to recover context or repeat unchanged native
  chain/reuse failures without a specific investigation hypothesis.
- Separate actual launcher form disposal from a deliberately retained test form.
  The latter must be constructed with a different workbook active before explicit
  target binding/closure. Do not remove the original refusal assertion.
- VBA80010007/80070057 and Excel nativec0000005/c0000028 are not desktop error5.
- VBE Debug/Reset after the earlier UOM modal preceded a native crash; avoid that
  intervention as a routine cleanup path. Diagnostic passes/forced exit are not fixes.
- Saved compilation, tracing/native observers and scoped COM-reference cleanup
  did not establish a cold native repair. Read batch-boundary evidence first.
- Do not read declared unused ListBox columns as if they were populated fields.
  Output movement's existing12-column loop still needs direct packaged proof;
  the snapshot error alone does not authorize a runtime fix.
- Read live probe logs with FileShare.ReadWrite; Get-Content/append sharing caused
  one logger stop. Never serialize PowerShell string ETS metadata into receipts.

## Assumptions to re-verify

Explicit resume; desktop cursor/input/capture availability; no Excel open; candidate
package pins/settings; both Git branches/status and unrelated-file pins. No desktop
monitor remains running. Current counters/old evidence are not a release percentage.

## Open questions and blockers

Protecting component RED is pending; partial-write probe insertion fails before
execution. Runtime Production coverage remains19/68 constructed controls,49 pending;
14 unconstructed legacy buttons and34 non-button handlers remain accounted for.
Other Operations/Admin coverage, human/NAS acceptance and current full Release1/
reusable Production proof remain open. Native failures and delayed exit have no
established root cause. The earlier6.5/10 estimate is not measured acceptance.

## Immediate next action

After explicit resume, repair/calibrate the partial-write insertion in
`Slice4beProductionComponentProbe.ps1`, verify the corrected snapshot, and run a
complete packaged component RED against the frozen extent candidate before runtime edits.

## Critical references

- Architecture/plan pointers: `0 plan docs/xlam_invSys/CURRENT_SPEC.md`,
  `expert guidance docs/CURRENT.md`; catalog16 D18 clause in `invSys-Design-v4.11.md`.
- `invSys-Controls-v1.md`, `invSys-Production-Tracking-Coverage-v1.md` in that spec directory.
- Code `tests/tooling/Test-Slice4beProductionComponents.ps1`,
  `Slice4beProductionComponentProbe.ps1`, `Slice4beProductionComponentActivity.ps1`.
- Source `src/Production/Forms/frmProduction.frm`: actual requirement/output Click
  handlers, WriteRequirementEditorToList, WriteOutputEditorToList, MoveSelectedListRow.
- Code `tests/integration/plan022_slice4be_remaining_acceptance.md`, component/UOM
  evidence above, and `plan022_slice4be_production_batch_boundary_results.md`.
