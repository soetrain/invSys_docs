# Saved-workbook restart behavior and remaining latency

## 1. Goal and release outcome

Complete Release 1 / Slice 4be under Architecture v4.11 and Plan 022. Goal active
and incomplete. This turn diagnoses native-window retention, corrects a test
expectation to preserve approved preference observations, and verifies saved-host
restart behavior. It does not complete visual or responsiveness acceptance.

## 2. Current verified state

- Last verified **2026-09-24 UTC**: code `main` **dd08366**, pushed. Docs `main`
  preceding commit **eb2a7b5**; this handoff accompanies controls **v1.231**.
- Runtime e986b51 / `deploy/validation-settings-diagnostic` unchanged: all
  **299 runtime/five package** pins verified. No VBA/form/Ribbon/XLAM edits.
- Final saved-host restart **34/34**, retaining all 27 old identities plus five
  saved-host and two exact Save boundaries; five instrumented compiles, normal
  unassisted closure, original settings restored, zero delayed Application failures.
- Guard calibration **32/32**. Static metrics unchanged:250components/6045procedures/
  132893lines,9literal/45unresolved calls,193duplicate groups;28size limits,
  three schemas and258PowerShell parses verified.
- Three final images directly reviewed but **not accepted as unobscured full
  controls**: Task Manager thumbnail overlays lower-right form area. Presentation
  content remains readable. Three clean earlier saved-host images are separate.
- **84.485-second** fresh paired-view Selected observation remains unexplained.
  Includes readiness observation and macro dispatch; internal attribution unproven.
- All test/controller/schema sessions terminal; **Excel closed**, no restoration
  state held. Preserve unrelated frmEventDetail.frm (+18 lines, Locked=True),
  modified docs handoff067 and untracked critique023. No other unfinished edits
  are intended after these commits.

## 3. Decisions and constraints

- Standing user authorization: **"you can always close Excel."** Preserve unsaved
  work, verify owned processes and distinguish assisted recovery from normal exit.
- D18 semantic inheritance remains approved; contradictions need explicit
  normative approval first. D8-A ordinary Auth read/provisioning and Event Detail
  Locked=False amendments remain unapproved. Carrier authority and transfer wire/
  provenance remain unresolved. Resume/full access never supplies approval.
- Preserve exact immutable System_Key, unknown columns, captured binding,
  headless authority, packaged reuse and every preceding GREEN identity.
- Normal R1 acceptance uses a saved `.xlsm`/`.xlsb`. Testing that condition does
  not authorize a runtime workaround, new hidden host workbook, or stress waiver.

## 4. Evidence and traceability

- Exact controls, failed calibrations and receipts:
  `tests/integration/plan022_slice4be_gui_resource_results.md` in code repository.
- Final controller `reports/runtime/guide-resource-diagnostic/087a6f7fdec84b63a3c416f3761d4ee7`;
  packaged root `reports/runtime/slice4be-viewer-published-read/83c161e3d78346debf143bac8cf76764`;
  receipt `reports/runtime/saved-workbook-restart-verification.json`.
  Interval00:39:27--00:43:41UTC; audit00:45:46UTC. Fresh boundary peak563GDI,
  maxsixXLMAIN; passive samples may include transient open-workbook windows.
- Native-window growth begins before pairing/capture. An Excel-only synthetic
  XLAM reproduces one retained window per VBA open/close:2->14 over12reads.
  COM reads stay2->2; initializing VBE still2->14; holding a separate saved
  workbook stays3->3. All four controls close normally and preserve source bytes.
- First add-in-only trace reaches82windows/3001peakGDI and stops; cleanup still
  grows, native close requests remain unprocessed and COM inspection stalls.
  Owned disposable process is terminated. Independent original settings snapshot
  is restored afterward. Full details distinguish this from later normal closures.
- Two saved-host31PASS/oneFAIL attempts expose only two required preference Save
  observations, not Config/prior-file mutation. D18 already requires this pair.
  The calibrated guard admits only exact correlated REQUESTED/COMPLETED records
  at Save, preserving prior hashes and both appended hashes through later reads.
- Earlier full chain32/48/15 and curation77/77+18images remain verified on the
  same frozen candidate; no unchanged reruns this turn. Earlier crashes and
  comparison333PASS/one Applied-label0x800AC472 exception remain recorded.

## 5. Do Not Repeat

- No unchanged add-in-only/broad comparison retry or OS-quota increase. Pure
  window retention is reproduced independently; its internal Excel cause is open.
- Do not suppress required preference observations or exclude the activity folder.
- Do not claim current unobscured captures, responsive operation, or full Slice
  acceptance from34/34. Do not relabel earlier clean images as final-run images.
- Do not compare or release an already released COM wrapper; even a PowerShell
  null comparison can throw. Use the existing Boolean fixture flag after closure.
- Keep original settings in controller memory until Excel exits. No reattachment/
  second Quit during passive lifecycle proving; assisted recovery stays explicit.
- Do not delete the older rejected temp directory ending
  `invsys-config-command-c365fc6fbfab48a79b8f842127ebf3fd`.

## 6. Assumptions to re-verify

Recheck Excel absence, frozen hashes, user changes and approvals. Process IDs are
historical. The slow selected-view boundary may relate to earlier COM rejection,
but this is unproven. Refresh/layout/activation recursion is a source lead only.

## 7. Open questions and blockers

All six workstreams remain in `plan022_slice4be_remaining_acceptance.md`:
comprehensive Production/Admin and remaining Settings/lifecycle coverage, pending
amendments, transfer, comparison/Viewer/guide gates and human/NAS acceptance.
Saved-host correctness is verified; full visible and responsiveness acceptance
remain open. Earlier guard/desktop/crash limitations retain their original scope.

## 8. Immediate next action

Instrument actual refresh/layout/activation entry counts around the packaged
selected-view observation in the saved-workbook fixture, retaining D18 validation
and obtaining unobscured captures, before proposing a runtime correction.

## 9. Critical references

- `src/Operations/Forms/frmActionPathView.frm`: RefreshView, UserForm_Layout,
  UserForm_Activate; `frmActionPaths.frm`: ReadViewEvaluation.
- `tests/tooling/Test-Slice4beGuideResourceDiagnostic.ps1 -SavedWorkbook`;
  `Slice4beGuideResourceTrace.ps1`, `Slice4beGuidePresentationRestart.ps1`.
- `Slice4bePreferenceSavePins.ps1` / `Test-Slice4bePreferenceSavePins.ps1`.
- `reports/runtime/guide-resource-static-final`; runtime/package pins unchanged.
- Authority pointers unchanged except this handoff pointer; controls v1.231.
