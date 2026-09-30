# UOM, draft and Settings restart continuation

## Goal and release outcome

Complete Release1 acceptance under Architecture v4.11 / Plan022, including Slice4be
comprehensive Event Viewer and Action Paths. **Goal active; release incomplete.**
This continuation retains packaged regressions and corrects an observed Settings
test-harness shutdown defect. It changes no runtime VBA or architectural contract.

## Current verified state

Last verified:2026-09-30,00:09 UTC. Both branches are `main`. Code **fe131c8** is
pushed; runtime remains **115b616**. Documentation **7a6db89** is the pre-handoff
head; this handoff/pointer and synchronized evidence form the following commit.
Code is clean. Preserve unrelated modified `last handoff/067 Partial Goal Action
Path NAS Contract.md` and untracked `expert guidance docs/023 Slice 4be Critique.md`;
both match their private pins.

Candidate **deploy/validation-production-components**, catalog16, remains
**unpromoted**. Controls1.277; coverage1.20; Production registration29/68,39 pending.
No Excel process or regression remains running. Before another gate, reverify
desktop availability and the five package pins in
`reports/runtime/production-component-controller/174213c4d7e44c9996606e6e42d205e4/package-pins.json`.
Never rebuild/deploy while relevant Excel files are open.

## Decisions and constraints

- Architecture D18 catalog16 was committed before runtime implementation under
  approved semantic inheritance. No new permission, saved-authority or inventory
  contract is introduced. A contradiction/material weakening still needs an
  approved normative decision before implementation.
- Keep exact immutable `System_Key`, unknown columns, captured-workbook/session
  binding, packaged launcher reuse, headless Core/Domain and all prior GREENs.
- The Settings restart helper closes all old workbooks, calls Quit, releases the
  application, then releases this isolated worker's completed-host COM references
  before the unchanged five-second wait. Replacement Excel is created afterward.
  The established reference-release helper's implementation is unchanged; its
  scope comment now explicitly includes completed hosts at a restart boundary.
- Fixed metadata records both reference counts and normal/forced exit. The trial's
  opt-in switch is removed: the final correction is the ordinary helper route.
  This is non-contract test tooling, not product behavioral RED or a native fix.
- User permits closing Excel; no reboot or lock-policy change is authorized. Keep
  invSys.StationUpdate disabled. On actual desktop error5, timestamp first failure
  and last success in UTC/Pacific, safely restore owned fixtures, pause the goal
  under the user's explicit stop condition, update handoff, commit/push and stop.

## Evidence and traceability

Exact controllers/results and qualifications are in code
`tests/integration/plan022_slice4be_production_component_results.md`.

| Gate | Current verified evidence |
| --- | --- |
| Component foundation retained |RED180/615 ->795/795; paired paths142/142; five builds/compiles/cold load, layout and static limits pass. Earlier continuation115 and the component evidence retain exact scope. |
| Combined visible UOM |**264/264**, exact prior identities, five compiles, normal closure/preservation; two principal captures reviewed. |
| Actual public UOM launcher |**61/61**, exact prior identities, five compiles, normal closure/preservation; three principal captures reviewed. Actual owner close/reopen retains the draft/custom column and REUSED outcome. |
| Draft/diagnostic paths |**390/390**, exact identities, five compiles, delayed normal closure/preservation. This regression requests no new principal captures. |
| Settings shutdown baseline |**202/202**; new receipt proves internal termination after five seconds. Final host exits unassisted. Earlier internal branches remain unobserved. |
| Settings reference control |**202/202**;10 references released, zero failures, internal restart unassisted within the unchanged wait; final host also exits normally. |
| Corrected ordinary Settings route |**202/202**, exact identities; again releases10 references without failures, both exits unassisted. Five compiles and preservation pass.00:03:54--00:08:11 UTC2026-09-30. |

All continuation gates record zero Excel Application1000/1001 events and preserve
settings/five packages. All310 PowerShell files parse; both diff checks pass.
Original workflow statements and reference-helper implementation are preserved.
No runtime source differs from115b616, so the existing compiled/static evidence
retains its scope:259 components/6078 procedures/133764 lines,9 literal/45 unresolved
calls,191 duplicate groups,28 limits held, form11735 lines; three schemas passed.

Ignored receipts: `production-components-uom-visible-verification.json`,
`production-components-uom-public-verification.json`,
`production-components-draft-verification.json`,
`production-components-settings-restart-observation.json`,
`production-components-settings-reference-control.json`, and
`production-components-settings-default-closure-verification.json`.
Default Settings result:
`slice4be-tracking-settings/0a0e512b5736471fa16bbf92bc1ca7af/green.json` under runtime
reports, with `preference-restart-closure.json` and reference-release sidecar.

Desktop checks remain clear through **00:09:44 UTC2026-09-30 /17:09:44 Pacific
2026-09-29**, continuing the overlapping observations from15:45 Pacific. No new
cursor/input/capture failures occur. Current finite observer is
`rdp-desktop-probe-d5fde564815349b38e164adf1e47eec2`, verified live, ending around
00:17:03 UTC. Use `rdp-desktop-probe-current.json` and reverify/renew before expiry.
Receipt: `production-components-regression-continuation-desktop.json`. This does
not prove the RDP display setting fixed locking or establish an idle timeout.

## Do Not Repeat

- Do not call old Settings runs unassisted: their internal branch was unobserved;
  the instrumented baseline explicitly used termination. Current default proof
  does not rewrite those records.
- Do not transfer this Settings harness correction to native chain/reusable
  failures without evidence. Earlier native diagnostics found no remaining live
  references and still needed termination; their cold c0000028 crashes occurred
  before cleanup. Read batch-boundary evidence before another unchanged variant.
- Do not repeat passing gates to recover context. Preserve exact prior check
  identities; a harness error is not meaningful product RED. Use shared-access
  FileStreams for live logs; avoid VBE Debug/Reset after uncertain modals.

## Assumptions to re-verify

Desktop/finite monitor availability, no Excel, package pins/settings, Git status
and unrelated-file pins. Goal status was re-read as active. No operational
deployment changed. The helper correction is bounded to isolated test ownership.

## Open questions and blockers

Current candidate still needs lifecycle615/615, native cancellation94/94, Settings
activity (previous468; verify catalog16 additions), smoke86/86 and applicable full
chain/live-role/reusable restart/export gates. Native failures remain unexplained.
The remaining39 Production buttons,14 legacy unconstructed buttons,34 non-button
handlers and other Operations/Admin coverage need complete accounting. Guide
transfer, broader comparison, human and applicable NAS acceptance remain open.
Use `plan022_slice4be_remaining_acceptance.md`; do not shrink the release goal.

## Immediate next action

Verify desktop/monitor, no Excel and candidate pins, then run
`reports/runtime/run-production-components-regression.ps1 -Mode Lifecycle`,
retaining all615 preceding identities and verifying normal closure/preservation.

## Critical references

- Architecture/plan pointers: `0 plan docs/xlam_invSys/CURRENT_SPEC.md` and
  `expert guidance docs/CURRENT.md`; controls/coverage catalogs beside the spec.
- Code component evidence and remaining-acceptance index.
- `Slice4beActionPathPreference.ps1`, `IsolatedAutomationCleanup.ps1` in tests/tooling.
- `plan022_slice4be_production_uom_activity_results.md` for frozen regression baselines.
- `plan022_slice4be_production_batch_boundary_results.md` and
  `plan022_slice4be_projection_boundary_results.md` for qualified native history.
