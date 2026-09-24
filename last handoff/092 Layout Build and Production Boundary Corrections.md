# Slice 4be: layout, build and Production boundary checkpoint

## 1. Goal and release outcome

Achieve Release 1 acceptance under Architecture v4.11/Plan 022 while preserving
accepted roles, identity and packaged behavior. Slice 4be remains **incomplete**;
the active goal remains active. This checkpoint closes bounded corrections and
retains the newly reproduced palette acceptance failure without weakening it.

## 2. Current verified state

Last verified: 2026-09-24 UTC. Both repositories are on main and pushed.
Code: **b215484** (Production primitive boundary), preceded by **bac31b6** (CRLF
test-region stripping) and **af6162c** (Action Path layout guard). Documentation:
**d2d5a09** before this handoff; controls version **1.234**. Normative/plan pointers
remain unchanged. Excel is closed; no test worker remains active.

Current bounded candidate: `deploy/validation-production-quiet-boundary`.
`deploy/current` was not replaced. Compared with
`deploy/validation-guide-layout-normalized`, only mProduction changes among 243
compiled components. All three five-package sets retain their hashes. The 299
original runtime pins differ only at frmActionPathView and mProduction.

Unrelated changes preserved: code `src/Operations/Forms/frmEventDetail.frm` (+18
measurement lines, still Locked=True); docs `last handoff/067 Partial Goal Action
Path NAS Contract.md` (+3/-3); untracked `expert guidance docs/023 Slice 4be
Critique.md`. EventDetail's +18 lines already exist in the frozen packages;
preserve them in builds without committing the user's source edit.

## 3. Decisions and constraints

The user authorizes closing Excel whenever needed; preserve unsaved work and
distinguish assisted shutdown from normal closure. One Excel gate at a time;
never build/deploy with relevant workbooks/add-ins open.

The layout guard avoids RefreshView on unchanged dimensions; genuine resize,
activation and view actions retain policy/integrity validation. The Production
correction uses the existing captured workbook-name bridge, with no new contract.
The CRLF correction only fixes detection of existing test-only markers.

Pending, not effective: D8-A ordinary Auth reads must not provision/repair/save;
Event Detail Locked=False for non-editable selection/scrolling. Earlier async
questions have no recorded answer. Carrier authority must follow warehouse Config
under D5; do not invent Windows-user authority or infer UOM permission. Guide
transfer's exact wire/provenance contract still needs specification. Semantic
inheritance permits discovered controls, not hidden architectural weakening.

## 4. Evidence and traceability

- Action Path: behavioral RED found two unwanted refreshes for unchanged size;
  GREEN **40/40**, including four sizes and all preceding identities. Captures
  disabled; latest selected-view observation 6.879 seconds is not full visible
  responsiveness acceptance. See `plan022_slice4be_layout_stability_results.md`.
- Build regions: RED **4 PASS/8 expected FAIL**, GREEN **12/12**; CRLF control
  builds/compiles five packages with all 243 components identical. See
  `plan022_crlf_build_regions_results.md`.
- Production boundary: actual two-batch form RED **50 PASS/two expected FAIL**;
  GREEN **52/52**, captured bridge entries and UI restoration both twice. Four
  instrumented compiles; normal closure/preservation; original static audit
  **14/14**. Five runtime packages compile; cold-start dependency check passes.
- Corrected candidate full chain **32/32**, ordered live roles **48/48**, Create
  Warehouse **15/15**, exact prior identities; normal unassisted closure,
  restored settings/reports, preserved hashes, zero Application failure events.
  Static metrics unchanged: 250 components/6045 procedures/132897 lines,
  9 literal+45 unresolved dynamic calls, 193 duplicate groups, 28 size limits;
  three schemas and 265 PowerShell parses pass.
- **Open palette failure:** the separate `ProductionReusable -ProductionRunOnly`
  gate is 0 PASS/one aggregate FAIL on corrected and preceding layout candidates.
  Both have EightPaletteRows=False; all 67 Boolean observations match. Reusable
  and Chai completion envelopes report OK, which does not waive the failed
  accepted eight-row requirement. This predates the boundary correction. The
  existing launcher harness may force its owned process closed after Quit;
  neither attempt establishes normal shutdown.

All Production details, exact report roots, candidate comparisons and intervals
are in `tests/integration/plan022_production_quiet_boundary_results.md`.
Root-cause hypothesis only: frmProduction.AddList assigns Height before disabling
IntegralHeight, possibly allowing MSForms to round the requested height. No
palette implementation change has been made. Required next proof is actual
packaged palette geometry at minimum/default, expanded and restored dimensions.

## 5. Do Not Repeat

Do not repeat broad reusable runs unchanged or remove EightPaletteRows to obtain
GREEN. Do not attribute its failure to the one-call boundary fix: the predecessor
control reproduces it. Do not conflate the base live-role 48 with ordered-chain
48: the chain replaces two Shipping checks with two Boxing checks; each gate's
exact identity set is separately verified.

Do not reuse rejected `deploy/validation-guide-layout` (archive lost EventDetail
measurement lines and retained test helpers). The accepted normalized candidate
and CRLF control comparison are separate. Do not retry native screenshots merely
because OpenInputDesktop succeeds; GetCursorPos previously still returned error5.
Overlay images and capture-disabled timings do not establish visible acceptance.

## 6. Assumptions to re-verify

Recheck desktop/input access, process state and candidate hashes before dependent
work. The last cursor probe at 01:23:59 UTC failed with error5 despite matching
Default desktop names; cause remains unknown. No evidence establishes throttling.
Palette height/setup-order causality is unconfirmed; obtain measured values before
choosing a correction. Historical acceptance is not proof of a new candidate.

## 7. Open questions and blockers

Restore the accepted palette layout; broader Slice 4be coverage and visible
acceptance remain open. Catalog11 has 62 controls; Production registers only UOM
Retrieve. Lifecycle observations require exact queued event/owning outcome
correlation, not SubmitReusableDesignEvent's Boolean/projected status or parsed
free text. Source discovery is retained at
`reports/runtime/production-observation-owner-discovery.md`.

Preserve the remaining checklist: pending Auth/Detail decisions, shared and Admin
coverage, guide transfer, Viewer/guide gates prepared at 98/355, comparison's
Applied pair-label failure, human comparison and applicable NAS acceptance.

## 8. Immediate next action

Add a focused packaged palette measurement through the actual Production form
path, record default/minimum, expanded and restored height/non-overlap RED against
the accepted eight-row requirement, then correct only the demonstrated cause.

## 9. Critical references

- `tests/integration/plan022_slice4be_remaining_acceptance.md` (complete gate index).
- The three bounded evidence files named above; Architecture v4.11, Plan022 and
  controls remain authoritative in their existing pointed locations.
- `frmProduction.AddList`, `TestRunListResponsiveLayoutReportForSize`,
  `TestReusableProductionRunActionContract` (EightPaletteRows assignment);
  `mProduction.BtnOpenProductionForm`; `tools/validate_plan022_packaged_launchers.ps1`.
- `reports/runtime/production-quiet-reusable-comparison.json`,
  `production-quiet-chain-verification.json`, `production-quiet-focused-verification.json`.
- `tests/tooling/Test-ProductionQuietBoundary.ps1` and
  `ProductionQuietBoundaryProbe.ps1` (disposable instrumentation; no runtime probe).
