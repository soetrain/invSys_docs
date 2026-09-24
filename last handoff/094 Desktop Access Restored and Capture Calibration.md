# Slice 4be: desktop access restored without elevation

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11 and Plan022, including
comprehensive Operations/Admin tracking and both Action Path presentations.
**Incomplete; goal active.** This checkpoint investigates the user's Windows
error5/admin-access question; no runtime or contract change is made.

## 2. Current verified state

Last verified: 2026-09-24 03:05:30 UTC. Both repositories main; code **66227da**,
docs **6f32af8** before this handoff. Controls **1.236**. Candidate remains
`deploy/validation-production-palette`; accepted deployment untouched. Excel
closed, both workers terminal. All five candidate hashes hold. Only the three
previously recorded runtime changes differ among 299 baseline pins.

Preserve unrelated changes: code `src/Operations/Forms/frmEventDetail.frm` +18
measurement lines, still Locked=True; docs handoff067 +3/-3 and untracked critique023.
Handoff093 and `plan022_production_palette_results.md` retain prior palette,
fixture, reusable Production and full-chain evidence; do not rerun them unchanged.

## 3. Decisions and constraints

User authorizes closing Excel; preserve unsaved work. User uses both physical
console and RDP and suggested administrator access. Read-only checks establish
the test shell and Excel are both Medium integrity, not elevated. No elevation,
ACL, security setting or desktop switch was performed. Keeping one desktop
connection active during visible tests is the practical recommendation; console/
RDP switching is only a hypothesis for the intermittent failure.

Pending, not effective: D8-A Auth read/provisioning separation and Event Detail
Locked=False for non-editable selection/scrolling. Prior questions remain without
recorded approval. D5 warehouse Config carrier authority and exact guide-transfer
wire/provenance specification still apply; no hybrid or implied approval.

## 4. Evidence and traceability

`tests/integration/plan022_slice4be_layout_stability_results.md` records this check.
At 03:01:20 UTC GetCursorPos succeeds, error0, WinSta0/Default matching input and
thread desktop. A single presentation attempt on the palette candidate records
**10 PASS/one harness exception** at Create guide fixture dispatch, before
restart/layout/captures. Five instrumented compiles and actual publication/source
journal fixtures pass. Normal unassisted closure and settings/package preservation.
Controller `guide-resource-diagnostic/ce214dc1673f4660b83d42812cd201ce`; report
`slice4be-viewer-published-read/1b8abd874a3940a4a54dcdbca9d6e15d` under reports/runtime.

Existing `Test-Slice4beCaptureForeground.ps1` then captures **all three cases**,
03:04:30--03:04:35 UTC; images individually reviewed, readable/unobscured. Blank
fixture only, normal closure, no elevation and no product acceptance claim.
Report `capture-foreground-calibration/41584aa806624030b24630a7e888a297`.
Receipt `reports/runtime/desktop-access-restored-verification.json`: zero
Application 1000/1001/1002 events at 03:05:30 UTC; preserved hashes and closed Excel.

## 5. Do Not Repeat

Do not declare permanent error5 resolution or require administrator access from
the old failure. Do not repeat the guide test unchanged: capture was never reached.
Capture calibration does not establish invSys visible acceptance. The exact
Create guide return/context is needed before diagnosing that fixture failure.
Do not weaken predicates or classify a fixture failure as D13 behavioral RED.

## 6. Assumptions to re-verify

Desktop access is transient: recheck before visible tests, particularly after
console/RDP changes. Candidate hashes and process state need checking before
dependent work. Last known approval state remains pending as above.

## 7. Open questions and blockers

Six workstreams remain in `plan022_slice4be_remaining_acceptance.md`. Catalog11
has 62 controls and Production only UOM Retrieve. Reachable Process/Recipe and
other Operations/Admin/shared handlers still need observation contracts/tests.
Lifecycle evidence must correlate exact owner events, not matching projection
status or free-text reports. Native visibility, guide transfer, broader Viewer/
comparison, human comparison and applicable NAS acceptance remain open.

## 8. Immediate next action

Resume the first reachable Production designer observation group: reconcile
precise local/owner outcomes with D18 and add packaged-handler RED before runtime
changes; inspect Create guide fixture dispatch separately before any visible retry.

## 9. Critical references

- Current specification/plan/controls pointers; handoff093 for preserved GREENs.
- Layout-stability evidence and remaining-acceptance checklist named above.
- `tests/tooling/Slice4beGuideRestartFixture.ps1:33`, FixtureControl and
  `modInventoryViewer.GuideDraftControlForTest`, frmActionPaths btnCreateGuide.
- Production discovery reports and modProductionUomAction/frmProduction handlers
  identified in handoff093; no tracking implementation started in this checkpoint.
