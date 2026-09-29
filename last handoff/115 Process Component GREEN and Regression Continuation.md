# Process component GREEN and regression continuation

## Goal and release outcome

Complete Release1 acceptance under Architecture v4.11 / Plan022, including Slice4be
comprehensive Event Viewer and Action Paths. **Incomplete; goal active.** The user
explicitly resumed after changing the RDP-client display timeout to Never. No new
desktop error5 occurs in this work period; stop again if it returns.

## Current verified state

Last verified:2026-09-29,23:37 UTC. Both branches are `main`. Code **2b3b768** is
pushed; runtime implementation is **115b616**, protecting RED checkpoint **8d616cd**.
Documentation **e510bd0** is the pre-handoff head; this handoff, pointer and scoped
regression updates form the following documentation commit. Code is clean. Preserve
the unrelated modified `last handoff/067 Partial Goal Action Path NAS Contract.md`
and untracked `expert guidance docs/023 Slice 4be Critique.md`; private pins match.

Current candidate: **deploy/validation-production-components**, catalog16,
**unpromoted**. Frozen extent candidate remains the RED baseline. No Excel process
remains and no regression is running. Never rebuild/deploy with relevant Excel
files open. Controls1.276; Production coverage1.19, registration **29/68**,39 pending.

## Decisions and constraints

- Architecture D18 catalog16 contract was committed in documentation **4dbbf76**
  before runtime work, under the approved semantic-inheritance rule. It adds ten
  requirement/output Add/Update/Remove/Up/Down observations. A contradiction or
  material weakening still requires an explicitly approved architecture decision.
- Typed owner checks captured workbook/context/current capability before local
  changes; loading/nested actions are suppressed. Existing Update identity/fallback/
  append, validation-time normalization, Remove reset, quantity/UOM and regulation
  semantics remain. Movement uses seven/ten owned fields. Recipe callers preserve
  their prior declared-column behavior through the extracted helper.
- Only STAGED concludes CommandCompleted; saved definitions remain Unchanged and
  empty source references never assert Domain application. Failed partial edits
  report uncertainty, not rollback. Optional tracking remains nonblocking.
- Preserve immutable exact `System_Key`, unknown columns, headless Core/Domain,
  packaged launcher reuse and every prior GREEN identity. User permits closing
  Excel; no reboot or lock-policy change is authorized. Keep StationUpdate disabled.
- On actual desktop cursor error5, record first failure/last success in UTC/Pacific,
  safely restore owned fixtures, pause the goal under the user's stop instruction,
  update handoff, commit/push and stop. VBA/native errors are not desktop error5.

## Evidence and traceability

Exact controllers/results and qualifications are in code
`tests/integration/plan022_slice4be_production_component_results.md`.

| Current candidate gate | Verified result |
| --- | --- |
| Protecting RED / focused GREEN |180 PASS/615 FAIL -> **795/795**, exact identities, five instrumented compiles, normal unassisted closure, settings/packages preserved; two principal captures reviewed. |
| Component publication/paired paths |**142/142**, all140 preceding checks plus two calibrated selected-field checks; six principal captures reviewed, normal closure/preservation. Ten actual actions match a distinct observed run in all three views. |
| Build/static |Five builds/compiles and cold Operations load pass. Exactly three changed plus three new compiled components,252 total, none removed. Static259 components/6078 procedures/133764 lines;9 literal/45 unresolved calls,191 duplicate groups unchanged;28 existing limits hold, form11736->11735. Three schemas pass;310 PowerShell scripts parse. |
| Layout |Three sizes/five pages, no bounds/overlap failures, native window actions pass; three principal captures reviewed, normal closure/preservation. |
| Settings / Instructions |**202/202 /411/411**, exact preceding identities, five compiles each, normal closure/preservation. |
| Instruction / UOM paths |**105/105 /84/84**, exact preceding identities, five compiles and six reviewed principal images per route; normal closure/preservation. Both conclusions explicitly avoid a Domain-application claim. |

Focused, path, layout and regression gates above record zero matching Excel
Application1000/1001 events.
Ignored receipts: `production-components-{focused,paths,layout,settings,instructions,instruction-paths,uom-paths}-verification.json`.
Build/static roots: `reports/runtime/production-components-build` and
`production-components-static`. Candidate pins are in
`production-component-controller/174213c4d7e44c9996606e6e42d205e4/package-pins.json`.

Desktop samples succeed from **22:45:27 to23:36:57 UTC /15:45:27 to16:36:57 Pacific**,
with overlapping three-second observers and no cursor/desktop/capture failures.
Receipt `reports/runtime/production-components-resumed-desktop-observation.json`.
Current finite observer `rdp-desktop-probe-07a9ada280174930802814d43e683eae` remains
active at this checkpoint, scheduled to end around23:41:23 UTC; reverify or renew
before screen work. Current pointer: `rdp-desktop-probe-current.json`.
The previous actual failure remains15:34:46.392 Pacific; no cause/idle timeout is
established and this success period does not prove the setting fixed it.

## Do Not Repeat

- Fixture defects are not product RED: VBE normalizes `.Text` casing; a no-argument
  procedure followed by a colon was parsed as a label. Explicit Call statements and
  Boolean ACTUAL setup calibration fixed the component adapters before valid RED.
- Component snapshot/movement must not treat unused declared ListBox slots as owned
  data. The final795/795 protects the actual seven/ten fields and partial failures.
- The first path capture selected REQUESTED. The next selected STAGED correctly but
  its getter read owner `Fields(1)`. Final142/142 calibrates and reads the visible
  field grid. Neither correction required runtime changes; earlier attempts remain.
- Preserve earlier native failures: Shipping Sent and cold reusable batch scale
  show Excel c0000028/RPC800706BE on earlier candidates. Tracing/VBE preparation,
  saved compilation, native observers and COM cleanup did not establish a repair.
  Read batch-boundary evidence before repeating an unchanged diagnostic variant.
- Do not use VBE Debug/Reset after an uncertain modal. Do not call a forced exit
  normal shutdown. Read live observer/worker logs with FileShare.ReadWrite.

## Assumptions to re-verify

Desktop access and finite observer health, no Excel open, candidate pins/settings,
both Git statuses and unrelated-file pins. Previous candidate results retain their
own scope. Goal-tool status was re-read as active during this turn; no new goal was
created and no completion or pause was requested.

## Open questions and blockers

Current candidate still needs combined visible UOM264/264, actual public UOM61/61,
draft/paths390/390, lifecycle615/615, native cancellation94/94, Settings observations
(previous468; expect added catalog16 exclusions, verify actual count), smoke86/86,
and applicable full-chain/live-role/reusable restart/export gates. Prior frozen
extent evidence is indexed in `plan022_slice4be_production_uom_activity_results.md`.
Native crash/delayed-exit causes remain unresolved. Remaining39 Production buttons,
14 unconstructed legacy buttons,34 non-button handlers and other Operations/Admin
coverage still need full accounting. Guide transfer contract, broader comparison,
human and applicable NAS acceptance remain open; do not shrink the release goal.

## Immediate next action

Verify desktop access/observer and no Excel, then run the existing combined visible
UOM regression against `deploy/validation-production-components` with
`Test-Slice4beProductionUomStaging.ps1 -Phase GREEN -CaptureEvidence -CheckActivity`,
retaining all264 preceding identities and verifying normal closure/preservation.

## Critical references

- Architecture/plan pointers: `0 plan docs/xlam_invSys/CURRENT_SPEC.md`,
  `expert guidance docs/CURRENT.md`; Architecture D18 catalog16 clause.
- Controls and Production tracking coverage catalogs beside the specification.
- Code component evidence and `plan022_slice4be_remaining_acceptance.md`.
- `plan022_slice4be_production_batch_boundary_results.md` and
  `plan022_slice4be_projection_boundary_results.md` for qualified native history.
- `Test-Slice4beProductionComponents.ps1`, `Test-Slice4beProductionInstructions.ps1`,
  `Test-Slice4beProductionUomStaging.ps1`; ignored regression wrapper
  `reports/runtime/run-production-components-regression.ps1` for other gates.
