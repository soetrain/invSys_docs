# Chain/curation GREEN and restart resource failure

## 1. Goal and release outcome

Complete Release 1 / Slice 4be under Architecture v4.11 and Plan 022. Goal active
and incomplete. This turn verifies the frozen candidate's chain and curation
gates, then identifies a concrete GUI-resource failure during presentation restart.

## 2. Current verified state

- Last verified 2026-09-23 23:58 UTC: code `main` **4a6bc96**, pushed. Earlier
  same-session/chain checkpoint **516f313** is pushed. Docs `main` preceding commit
  **89bd554**; this handoff accompanies controls v1.230 and the next docs commit.
- Runtime e986b51 and `deploy/validation-settings-diagnostic` unchanged: 299
  runtime/five package pins. Preserve unrelated frmEventDetail.frm (+18 measurement
  lines, Locked=True), modified docs handoff067 and untracked critique023.
- Same-session projection cut **35/35**, exact prior ordered prefix, normal
  closure. Complete chain **32/32**, live roles **48/48**, Create Warehouse **15/15**,
  every prior identity retained, normal unassisted closure, zero delayed Application
  failures, settings/three reports restored, 299 runtime/253 tooling/five package
  hashes preserved. Earlier combase.dll/ntdll.dll crashes remain unexplained.
- Curation **77/77**, 18 directly reviewed images, five instrumented compiles,
  immediate unassisted closure, zero Application failures, 299 runtime/195 test/
  five package hashes preserved. Library captures prove layout/direct entry with
  unavailable recordings, not populated-recording acceptance.
- Presentation restart **25 PASS/two harness failures**, two prior checks unreached.
  Exact preference/pairing assertions pass, but a captured insufficient-memory
  dialog blocks restored Compare. Close macro fails with 0x800A03EC; fixture
  cleanup fails. Two captures accepted, one rejected, one diagnostic duplicate.
- All current sessions terminal. **Excel closed with assistance** after this
  failure. No live controller or pending in-memory restoration state remains.

## 3. Decisions and constraints

- Standing user authorization: **"you can always close Excel."** Preserve unsaved
  work and record assistance honestly. Do not ask again for routine Excel closure.
- D18 semantic inheritance approved. New controls need exact normative/Plan/
  catalog entries and D13; architecture contradictions require explicit approval.
- D8-A ordinary Auth read/provisioning and Event Detail Locked=False amendments
  remain unapproved. Carrier authority and exact guide-transfer wire/provenance
  remain unresolved. Resume/full-access messages do not approve those amendments.
- Preserve immutable exact System_Key, unknown columns, captured workbooks,
  headless authority, packaged reuse and all prior GREEN identities. One Excel
  gate at a time; no build/deploy while relevant workbooks/add-ins are open.

## 4. Evidence and traceability

- Projection/chain: `tests/integration/plan022_slice4be_shutdown_header_results.md`.
  Focused root `projection-live-control/5892044845554763bbd156cefc344971`;
  receipt `settings-diagnostic-projection-checked-chain-verification.json`.
  Chain interval23:28:23--23:33:45 UTC; delayed audit23:34:07 confirms zero events.
- Curation root `slice4be-viewer-published-read/5ef27faf08dd4ca7a8c78dcde636261b`;
  receipt `settings-diagnostic-fixed-regression-curation-verification.json`.
  Detailed scope: `plan022_slice4be_guide_action_curation_results.md`.
- Restart root `slice4be-viewer-published-read/4d9d2c76d9f14bba9afc3b815f3f2e3f`;
  receipt `settings-diagnostic-fixed-regression-restart-attempt-verification.json`.
  Detailed failure: `plan022_slice4be_guide_presentation_results.md`.
- Failed process has **9,991 current/10,000 peak GDI objects**, configured quota
  **10,000**; USER9,875/quota10,000. Native enumeration counts **301 XLMAIN** windows,
  but later COM recovery inspection has **zero workbooks/zero workbook windows**.
  GDI quota was reached; the code retaining/creating windows remains unproven.
  Do not infer physical RAM shortage, Codex throttling, or the cause of earlier
  desktop error5 from this finding. Source facts are local and redacted.
- Recovery acknowledges only the exact owned memory dialog. Quit on the empty
  identified process returns but leaves it alive; it is then terminated after
  identity and zero-workbook/window checks. No operational workbook is changed.
  The original harness executes registry restoration before exit. Later recovery
  preserves its post-harness settings snapshot, but independent equivalence to
  the original pre-test settings is **unproven**. Partially cleaned generated
  fixture remnants are preserved; do not guess their ownership for deletion.
- `paired-restart-{memory-resource-facts,gui-quotas,window-counts,recovery-workbook-facts,assisted-closure}.json`
  under `reports/runtime/` preserve exact facts. Audit23:55:42 has zero Application
  events; visible memory failure/assisted recovery still fail acceptance.
- Prior static evidence remains unchanged:250components/6045procedures/132893lines,
  9literal/45unresolved dynamic calls,193duplicate groups; three schemas/28limits.
  The new chain regenerates its static ratchet, and all253PowerShell files parse.

## 5. Do Not Repeat

- Do not rerun the unchanged chain or curation: their current-candidate gates pass.
- No broad restart/comparison retry until resource growth is localized. Do not
  increase OS quotas or automatically add command retries to conceal this failure.
- Inspect resource growth around actual pairing/rendering **and native foreground
  capture separately**. Existing evidence does not identify which causes growth.
- Never mark the restored Compare image acceptable; its memory dialog is material.
- Use local-time Get-WinEvent bounds. Zero Application events/exit0 alone is not
  clean shutdown. No COM reattachment/second Quit during passive observation;
  after a terminal failure, authorized assisted recovery must remain explicit.
- Do not delete the previously rejected temp directory ending
  `invsys-config-command-c365fc6fbfab48a79b8f842127ebf3fd`.

## 6. Assumptions to Re-verify

Check Excel absence, frozen hashes and approvals. All recorded process IDs are
historical. Core source unchanged; no runtime fix or new architecture decision
has been made for GUI-resource growth. Do not treat the earlier header-scan fix
as a general Excel resource-leak correction.

## 7. Open questions and blockers

Comprehensive Production/Admin coverage, remaining Settings/lifecycle controls,
guide transfer, pending amendments, restart resource failure, comparison Applied
label0x800AC472, Viewer/guide candidate gates and human/NAS acceptance remain open.
The restart resource failure may be related to other UI symptoms, but that is an
unproven hypothesis. Existing Settings780/202, owner460, UOM228, Boxing1707PASS/
seven pendingD8-AFAIL, reader124 and comparison333PASS/one exception retain their
original scope/closure limitations; see the remaining acceptance checklist.

## 8. Immediate next action

Prepare a bounded fresh-session paired-view diagnostic with native GDI/USER and
XLMAIN counts before/after each actual action and capture, preserving settings
restoration in a controller until Excel is closed, before changing runtime code.

## 9. Critical references

- Authority pointers unchanged; controls v1.230; six-workstream checklist:
  `tests/integration/plan022_slice4be_remaining_acceptance.md`.
- `tests/tooling/Slice4beGuidePresentationRestart.ps1`: RestartPair, RestartCapture.
- `tests/tooling/Test-Slice4beConfigCommands.ps1`: Run, CaptureOwnedFormEvidence,
  finally/registry restoration. No per-action timing currently proves the slow call.
- `src/Operations/Forms/frmActionPathView.frm`: RefreshView/UserForm_Layout;
  source inspection is a lead only, not established root cause.
- `reports/runtime/run-settings-diagnostic-comparison-fixed.ps1`: curation/restart
  prefixes already used; never overwrite. Viewer/guide remain unused, expected
  98/355 checks and3/52images after reader-boundary additions.
