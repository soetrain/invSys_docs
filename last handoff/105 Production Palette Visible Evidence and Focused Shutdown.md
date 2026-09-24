# Slice 4be: Production palette capture and focused shutdown

## 1. Goal and release outcome

Complete Release1/Slice4be under Architecture v4.11 and Plan022. Goal remains
active/incomplete. This checkpoint advances 4be.6 preservation evidence for the
existing eight-row Production palette, without runtime or contract changes.

## 2. Current verified state

Last verified: 2026-09-24, after22:00 UTC. Both repositories are on main; prior
code checkpoint is **3f1f499**, committed/pushed; docs parent is70682e0. The code
checkpoint contains palette capture/cleanup tooling and evidence; the ensuing docs checkpoint contains
Architecture evidence, Plan022, controls1.247 and this handoff. Verify final heads
and upstream equality before continuing. Preserve unrelated modified handoff067
and untracked critique023; both match their private pins. Runtime evidence is ignored.

Frozen candidate remains `deploy/validation-detail-multiline-labels`; all five
package hashes match prior pins. No rebuild is required for this tooling-only
checkpoint. Excel and test controllers are closed; settings are restored.

## 3. Decisions and constraints

No architectural decision changes. D8-A and D18 scrolling remain approved.
Palette rows in the new capture fixture are explicitly display-only: no synthetic
inventory identity or business action is introduced. Preserve every prior gate.
The corrected cleanup applies only to ProductionPaletteProbe, not broader reusable
Production or other launcher branches. Close uses the actual form handler and
checks disposable workbook roots before closing fixture books and add-ins.

## 4. Evidence and traceability

Primary record: `tests/integration/plan022_production_palette_results.md`.
Runs below are under `reports/runtime/production-palette/`:

- Captured predecessor RED: `fb932e5d6c6a4531a3dd5f91f014ec52`, **12 PASS/four
  expected height FAIL**, including shutdown. Current GREEN:
  `495a545ed9224aacb9749671ea9c618f`, **16/16**, retains all11 prior identities.
  UTC21:55:12--21:56:17; four instrumented compiles per gate, settings/packages
  preserved, normal unassisted closure, zero Application failure events.
- Four GREEN images reviewed: eight complete rows at minimum/default/native
  maximize/restored. Predecessor default also shows eight short fixture rows:
  height RED proves requested-height loss, not clipping of row eight.
- Separate cleanup RED: `193f0899a05f478594e9d56bc8e110fe`,11 PASS/one expected
  shutdown FAIL. GREEN: `12e1a51df94e426a938844770d238e29`,12/12. Exact owned
  process termination is explicitly observed rather than inferred from absence.
- Static regeneration: `production-palette-visible-static`; verification receipt
  `production-palette-visible-static-verification.json`. Runtime metrics unchanged:
  251 components,6050 procedures,133033 lines,9 literal/45 unresolved calls,
  191 duplicate groups and28 oversized limits. Three schemas and279 PowerShell
  parses pass. No runtime source changed.
- Earlier five-package build/compiles, Detail88/88 with14 images, Viewer98/98
  with3 images, chain/live-role/Create Warehouse32/48/15 remain applicable at
  their recorded scope in `plan022_slice4be_multiline_results.md`.

**Correction to handoff104:** smoke86/86 is behavioral GREEN, but its normal-
closure claim was unsupported. The helper may terminate Excel after one second
and did not record that branch. The old private receipt is preserved and explicitly
superseded by `detail-multiline-packaged-closure-scope-correction.json` and the
updated multiline record. Separate Detail/Viewer/chain unassisted closure holds.

## 5. Do Not Repeat

Do not equate a populated list or geometric predicate with reviewed visibility.
Do not infer normal exit from Excel absence where cleanup may terminate it.
Do not rerun the complete Shipping/Boxing route without addressing its preserved
caption/z-order failure. Do not edit pinned inputs during gates or attach a second
COM controller during cleanup. Earlier failed capture and shutdown attempts remain.

## 6. Assumptions to re-verify

Error5 recurred at21:49 and cursor access returned at21:52/21:55 without policy
edits. Read-only selected values: local machine inactivity0, selected RDP timeouts
absent, screen-saver activation1 without established timeout/locking policy.
These do not establish cause or the RDP client state. Correct private policy file:
`desktop-idle-policy-verified.json`; the first extraction retained stale values
after missing properties and is explicitly named `desktop-idle-policy-observation-invalid.json`.
Never use that invalid file as evidence. Recheck desktop access and Excel closure.

## 7. Open questions and blockers

Palette captures reveal **Committed / Used** wrapping into a clipped second
line in `hdrManagerCheck8`. AddColumnHeaders gives a14-point height and68-point
width for the72-point column; this is a suspected fit cause, not yet packaged RED.
The maximized capture also has a small tooltip over lower status. Palette is
unobstructed; whole-form/human acceptance remains open.

The maintained remaining checklist retains comprehensive tracking (61 pending
Production buttons), guide transfer/provenance, comparison Applied-pair read,
presentation/restart, carrier D5 authority, seven owner-image corrections, broader
Shipping/Boxing capture, Settings shutdown, reusable cleanup, human/NAS acceptance.
Do not substitute this bounded checkpoint for the Release1 objective.

## 8. Immediate next action

Add a packaged protecting test for complete `hdrManagerCheck8` caption fit at
supported Production sizes and establish behavioral RED on the frozen candidate
before correcting the visible clipping under the existing header-readability contract.

## 9. Critical references

- Architecture v4.11; Plan022; controls1.247 and remaining-acceptance checklist.
- `tests/tooling/ProductionPaletteProbe.ps1`, `Test-ProductionPaletteLayout.ps1`.
- `tools/validate_plan022_packaged_launchers.ps1`.
- `src/Production/Forms/frmProduction.frm`: AddColumnHeaders, RUN_CHECK_WIDTHS,
  hdrManagerCheck8, mBtnClose_Click.
- `tests/tooling/Test-Plan022Slice4ayProductionRunListLayout.ps1` and
  `Test-Slice9ProductionLayout.ps1` are supplemental source checks.
- `reports/runtime/production-palette-visible-verification.json` records image
  hashes, identities, settings/package/user-file preservation and event audit.
