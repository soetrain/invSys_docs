# Slice 4be: multiline Event Detail GREEN

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11 and Plan022 without
regressing accepted workflows. Goal remains active/incomplete. The Slice4be.3
multiline refinement now has packaged native-input and reviewed visible evidence.
Comprehensive tracking, broader gates and human acceptance remain open.

## 2. Current verified state

Last verified: 2026-09-24, after21:39 UTC. Both repositories are on main.
Code **d0581c7** is committed/pushed. Docs parent **b8dddf6**; the ensuing commit
contains the normative refinement, Plan022, controls1.246 and this handoff.
Verify its push/upstream state. Preserve unrelated handoff067 edits and untracked
critique023; their private reference hashes still match. Runtime evidence is ignored.

Frozen candidate: `deploy/validation-detail-multiline-labels`. Five packages
build/compile; Operations cold-start resolves locally. Of244 compiled components,
only Operations/frmEventDetail differs from validation-detail-original-text.
Accepted deployment and both predecessors are preserved. Excel and all test
controllers are closed; local settings and tracked reports are restored.

## 3. Decisions and constraints

D8-A and D18 scrolling remain approved. Before runtime editing, the normative
specification, Plan and catalog named the multiline Frame/Labels under the user's
approved semantic-inheritance rule. This implements existing complete-text
reachability without an editor, extra selection, authority access or activity.
Office Label captions normalize CR/LF separators for display; original loaded
strings remain binary-exact and no source value is normalized or written back.
Only already permitted fields are rendered, in profile order, with context clearing.

## 4. Evidence and traceability

Primary sanitized record: `tests/integration/plan022_slice4be_multiline_results.md`.

- Final captured RED **57 PASS/31 expected FAIL** -> **88/88 GREEN**, exact
  identities and all56 preceding checks retained, including seven native-input
  checks. UTC21:30:00--21:34:24. Five instrumented compiles per run, normal closure,
  preserved settings/packages/tooling and zero Application failure events.
- Fourteen principal Detail images reviewed: mixed separators, literal escapes,
  Unicode/tab/ampersand, last of multiple fields and wide final line through END.
- Viewer/filter/Shipping-state **98/98**, exact prior checks, three reviewed
  images, normal closure/settings/pins/events. UTC21:35:06--21:38:08.
- Packaged smoke **86/86**, exact prior checks and normal closure/settings/report/
  package preservation; zero Application failures. UTC21:38:45--21:39:06.
- Full chain/live-role/Create Warehouse **32/48/15**, exact prior identities,
  normal closure, settings/three reports restored, pins preserved during gate,
  zero Application failures. UTC21:16:16--21:21:48.
- Static:251 components,6050 procedures,133033 lines (+3 small helpers/+47 lines),
  nine literal/45 unresolved calls,191 duplicate groups; all28 oversized limits
  preserved. Three schemas, layout8/8 and7/7,278 PowerShell parses pass.

Private receipts: `detail-multiline-dpi-visible-verification.json`,
`detail-multiline-visible-image-review.json`, `detail-multiline-viewer-visible-verification.json`,
`detail-multiline-packaged-verification.json`, `detail-multiline-labels-chain-verification.json`,
`detail-multiline-dpi-static-verification.json` under reports/runtime.

## 5. Do Not Repeat

MSForms Label.UseMnemonic is unsupported at runtime despite successful compile;
use the actual Accelerator property. The rejected candidate/438 attempt is retained.
UserForm.Controls includes nested controls; validate their actual parent bounds.
Do not compare DPI-virtualized coordinates directly with physical screenshot
pixels. The detector correction has independent **4 PASS/4 FAIL ->8/8** synthetic
calibration at100/150/200%, and still rejects the preserved locked-scroll pair.
The initial captured87/1 attempt is retained, not relabelled as GREEN.

Do not blindly repeat full Shipping/Boxing: its1632/1 caption/z-order failure and
82 unreached checks remain. Focused multiline success does not resolve that defect.
Do not edit pinned inputs while tests run, hash open XLAMs, reattach COM during
original cleanup or issue a second Quit. Preserve original settings snapshots.

## 6. Assumptions to re-verify

Desktop cursor access failed with error5 earlier, then returned at21:17 and21:23;
all final captured gates succeeded. Root cause is unproven. Recheck availability
and Excel closure before the next gate. Earlier intermittent native Excel crashes
remain unexplained. Production palette native evidence still has not been rerun.

## 7. Open questions and blockers

Use the maintained remaining-acceptance checklist. Comprehensive Operations/Admin
tracking (61 pending Production buttons), guide transfer/provenance, comparison/
presentation, carrier D5 authority, seven owner-image corrections, broader capture/
normal shutdown, reusable Production visible evidence and human/NAS acceptance
remain. This checkpoint closes only scoped automated multiline acceptance.

## 8. Immediate next action

Add guarded visible capture to the focused packaged Production palette gate and
verify its accepted eight-row behavior on the frozen current candidate now that
desktop input works, preserving existing launcher/reuse checks and honest cleanup evidence.

## 9. Critical references

- Architecture v4.11 D18/D8-A; Plan022; controls1.246.
- `tests/integration/plan022_slice4be_remaining_acceptance.md`.
- `tests/integration/plan022_slice4be_multiline_results.md`.
- `tests/integration/plan022_production_palette_results.md`.
- `tests/integration/plan022_slice4be_capture_lifecycle_results.md`.
- `tests/tooling/Test-ProductionPaletteLayout.ps1`, `ProductionPaletteProbe.ps1`.
- `tools/validate_plan022_packaged_launchers.ps1`.
- `tests/tooling/Slice4beViewerEventDetail.ps1`, `Slice4beDetailMultiline.ps1`.
- `tests/tooling/Slice4beDetailScrollEvidence.ps1`, `Test-Slice4beDetailScrollEvidence.ps1`.
- `src/Operations/Forms/frmEventDetail.frm`: RefreshFields, ClearMultiline,
  RenderMultiline, AddMultilineLabel, ClearContent.
