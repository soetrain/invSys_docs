# Slice 4be: Production palette GREEN; tracking coverage next

## 1. Goal and release outcome

Finish Release 1 acceptance under Architecture v4.11/Plan022, including
comprehensive Operations/Admin Event Viewer and Action Path coverage. The active
goal and Slice 4be remain **incomplete**. This turn restores the accepted
Production palette predicate and preserves reusable Production/release behavior.

## 2. Current verified state

Last verified: 2026-09-24 UTC. Both repositories are main and pushed. Code
**93cde9b**; docs **19bdb3b** before this handoff; controls version **1.235**.
Excel is closed and all workers are terminal. Candidate:
`deploy/validation-production-palette`; `deploy/current` remains unchanged.
Only frmProduction differs among 243 compiled components from the preceding
`validation-production-quiet-boundary` candidate. Four five-package sets retain
their hashes. Of 299 original runtime pins, only frmActionPathView, mProduction
and frmProduction differ, as expected from the three bounded corrections.

Unrelated changes preserved: `src/Operations/Forms/frmEventDetail.frm` (+18
measurement lines, Locked=True; those lines already exist in frozen packages);
docs `last handoff/067 Partial Goal Action Path NAS Contract.md` (+3/-3); untracked
`expert guidance docs/023 Slice 4be Critique.md`. Do not commit or discard them.

## 3. Decisions and constraints

The user authorizes closing Excel; preserve unsaved work. One Excel gate at a
time, no build/deploy while relevant workbooks/add-ins are open. Do not equate
assisted shutdown with normal closure or automated geometry with native visibility.

The runtime change moves IntegralHeight=False before Height assignment in
frmProduction.AddList. It honors existing requested dimensions; no contract or
test threshold changes. The shared New-AuthWorkbook fixture now formats its
named hash column as text. This is test setup only, not a runtime Auth amendment.

Pending, not effective: D8-A ordinary Auth reads must not provision/repair/save;
Event Detail Locked=False for non-editable selection/scrolling. Prior async
questions have no recorded answer. Carrier authority remains warehouse Config
under D5; exact guide-transfer wire/provenance details still need specification.
Semantic inheritance permits discovered controls, not architectural weakening.

## 4. Evidence and traceability

Exact evidence: `tests/integration/plan022_production_palette_results.md`.

- Palette packaged RED **7 PASS/four expected FAIL**: requested 96 points became
  88.45 at minimum/default/restored sizes and in the real factory; reassignment
  after IntegralHeight=False preserved 96. Cause is now confirmed.
- GREEN **11/11**: minimum/default/restored/factory 96; expanded 118.6. Every RED
  identity retained; all-page bounds/non-overlap and original launcher pass.
- Five-package build/explicit compile/cold-start dependency checks pass. Source
  Run List/layout checks **7/7 and 8/8**; static metrics unchanged at
  250 components/6045 procedures/132897 lines, 9+45 dynamic calls, 193 duplicate
  groups, 28 size limits; three schemas and 268 PowerShell parses pass.
- First broader attempt fails fixture sign-in before callbacks. A separately
  reproduced opaque-value coercion defect gives fixture RED **4 PASS/eight FAIL**,
  then **12/12** after text formatting. Values are never reported; private test
  workbook files are removed. The original sign-in failure's precise cause is
  not proven. New failure reporting is sanitized numeric/code classification.
- Corrected-fixture reusable gate passes **one aggregate**. All 67 Boolean
  observation identities remain; only EightPaletteRows changes False -> True.
  Two batches and Chai fork/convergence retain prior successful behavior. Its
  existing launcher helper may terminate Excel after Quit, so no normal-closure
  claim is made for this gate.
- Final candidate chain **32/32**, ordered live roles **48/48**, Create Warehouse
  **15/15**, exact preceding identities; normal unassisted closure, restored
  settings/three reports, five package/266 tooling pins preserved, zero Application
  failures at 02:52:02 UTC. Interval 02:44:45--02:51:13 UTC.

Prior Action Path 40/40, build-region 12/12 and Production boundary 52/52/14/14
remain recorded at their own scopes; no unchanged broad rerun is needed.

## 5. Do Not Repeat

Do not remove/weaken EightPaletteRows or retry the former palette failure: it is
fixed and the broader predicate is GREEN. Do not rerun the failed generic fixture
unchanged; use the corrected shared helper. Do not log generated secrets/hashes
or infer the original missing Auth status from later success. Supply RepoRoot
explicitly to Test-Slice9ProductionLayout.ps1; its default-path failure is harness
failure, not product RED. Do not rebuild accepted packages simply to refresh dates.

## 6. Assumptions to re-verify

Desktop preflight at 02:35:20 UTC still returns GetCursorPos=False/error5; no
capture was attempted. Cause remains unknown and no throttling claim is proven.
Recheck input access only before dependent visible work; do not repeat full
capture attempts without changed access. Confirm processes/candidate hashes and
pending approvals before dependent work.

## 7. Open questions and blockers

See `tests/integration/plan022_slice4be_remaining_acceptance.md` for the complete
six-workstream scope. Catalog11 has 62 controls; Production registers only UOM
Retrieve. Reachable Process/Recipe, remaining Operations/Admin/shared controls
still need D18 classification, catalog entries and actual-handler tests.
Lifecycle diagnostics must correlate exact queued events with owner outcomes;
SubmitReusableDesignEvent's Boolean/projected status is insufficient evidence of
that event's application. Do not parse free-text reports into new authority.

Still open: guide transfer, broader Viewer/guide/comparison gates, the pending
Auth/Detail decisions, comprehensive coverage, human comparison and applicable
NAS acceptance. Native visible evidence remains separate from automated GREEN.

## 8. Immediate next action

Reconcile the first reachable Production designer observation group against D18,
record its exact catalog/plan/control contract, then add packaged-handler RED for
missing observations and stale-context refusal before implementing tracking.

## 9. Critical references

- Architecture v4.11 / Plan022 / controls via current pointers; no pointer changes.
- `plan022_production_palette_results.md` and the remaining checklist above.
- `reports/runtime/production-palette-{focused,reusable,chain}-verification.json`;
  `production-palette-desktop-facts.json`; reproduction commands in evidence.
- `reports/runtime/production-observation-owner-discovery.md` and
  `production-observation-surface-discovery.json` (source-only discovery).
- frmProduction Process/Recipe handlers; modProductionReusableDesigns;
  modProductionUomAction; Core modActivityCatalog/modActivityReferences.
- `Test-ProductionPaletteLayout.ps1`, `ProductionPaletteProbe.ps1`,
  `Test-AuthFixtureText.ps1` (test instrumentation remains outside runtime).
