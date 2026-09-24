# Slice 4be: Event Detail scrolling GREEN and Production visible evidence

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11 and Plan 022, preserving
accepted workflows. Goal is active and incomplete. This checkpoint implements
the approved D18 single-line scrolling correction and advances packaged/visible
regression evidence; approved D8-A Auth separation remains unimplemented.

## 2. Current verified state

Last verified: 2026-09-24, after 17:34 UTC. Both repositories are on main. Code
commit **c1f0fba** contains this implementation/evidence; documentation base is
**03ad0a6**, with controls/Plan updates and this handoff in the ensuing commit.
Both commits must be pushed and upstream status checked before ending the turn.

Frozen candidate: `deploy/validation-approved-detail-scroll`. All five packages
build/compile and Operations cold start passes; only frmEventDetail differs among
244 compiled components from `validation-production-diagnostics`. The previously
pending 18-line font-measurement helper is included unchanged with the approved
Locked=False change, so committed form matches the tested candidate. Its own
earlier RED/GREEN is retained. Accepted deployment is unchanged. Controls1.241.

Excel is closed; every controller has terminated and settings are restored.
Preserve unrelated docs changes: handoff067 and untracked critique023. Runtime
reports/screenshots remain ignored and private. No implementation edits remain
uncommitted after the checkpoint commits.

## 3. Decisions and constraints

Do not ask again for D8-A or Event Detail approval. Both were explicitly approved
2026-09-24 and are normative. Multiline readability remains separately open.
Locked=False enables only non-editable selection/scrolling; no workflow handler
was added. D12/D13, exact System_Key, unknown columns and captured binding remain.

For D8-A, preserve explicit Generate/Create Warehouse and station provisioning.
Read the existing D2 operator sign-in rule8 before writing blanket sign-in
preservation assertions: it expressly authorizes a credential-validated, same-user
S1-to-current-computer capability transition. The source performs that transition
after credential validation in TryTransitionLegacyS1CapabilitiesForCurrentComputer.
Distinguish ordinary reads from that established authorized transition; do not
silently remove it or invent an exception. Any unresolved normative conflict must
be surfaced before implementing the affected behavior.

## 4. Evidence and traceability

- Event Detail: fresh RED **43 PASS/one expected movement FAIL**, then **44/44**
  GREEN, every prior identity retained. Native typing preserves exact values.
  Four default/scroll/maximized/restored images reviewed; source bytes/read guards
  pass. See `plan022_slice4be_detail_overflow_results.md` for exact reports.
- Viewer/filter/Shipping-state: **98/98**, retaining 94 historical identities;
  three reviewed images, immediate normal closure and zero Application failures.
- Owner: **460/460**, every prior identity, five instrumented compiles, delayed
  normal closure, zero Application failures. Of 23 reviewed images, 16 are clean,
  four have taskbar overlays, three reset-dialog images are wrong Settings crops.
  Seven remain limited/rejected visual evidence despite passing capture predicates.
- Production: **390/390**, every prior identity, five instrumented compiles and
  ten reviewed designer/Detail/diagnostic images. Normal delayed closure, restored
  settings, preserved packages and zero Application failures. Final controller
  `production-designer-controller/67960ea6bc4944cdbf47c18e7a34a133`; report
  `slice4be-production-designer/54878f9a2b874b10878c6d4b1f6a0491/green.json`.
  See `plan022_slice4be_production_paths_results.md` for two retained partial runs.
- Full chain/live roles/Create Warehouse: **32/48/15**, exact prior checks,
  normal closure, settings/three tracked reports restored, five packages and 273
  tooling hashes preserved, zero Application events 1000/1001/1002.
- Static: 251 components, 6050 procedures, 133025 lines, 9 literal/45 unresolved
  dynamic calls, 193 duplicate-body groups and 28 oversized limits unchanged.
  Source layout 8/8 and 7/7; 273 PowerShell files parse. Offline calibrations:
  readiness 44/44, guide-state retry 36/36, final failure diagnostics 83/83.

Private receipts use `reports/runtime/approved-detail-scroll-` prefixes;
`-production-ready-verification.json`, `-chain-verification.json`,
`-owner-verification.json`, `-viewer-verification.json`, `-focused-verification.json`
and `-static-verification.json` preserve their individual scopes. They are not
publication artifacts. Do not infer full Slice4be acceptance from these gates.

## 5. Do Not Repeat

Do not restart completed broad gates without a relevant change. Production's
first current-candidate capture stopped at 379 PASS/one caption-activation exception;
the second at 340 PASS/one 0x800AC472 boundary failure. Neither is GREEN. The first
failure cause and the exact second failed action are unproven. Final success uses
the existing bounded pre-dispatch readiness check; commands are never replayed.

No COM reattachment or second Quit during post-Quit cleanup. Keep the original
controller alive to restore its in-memory settings snapshot; normal exit took time.
Review images individually: a saved PNG or nonblank predicate is insufficient.
Do not commit runtime screenshots or sensitive generated fixtures/reports.

## 6. Assumptions to re-verify

Desktop access worked throughout these scoped gates without elevation, following
the user's host-setting changes. Earlier error5 cause is still unproven. Verify
Excel closure and candidate hashes before another gate. Run one Excel gate at a
time; do not edit pinned tooling mid-run or rebuild an open package.

## 7. Open questions and blockers

D8-A focused packaged Core RED/GREEN; Event Detail multiline; comprehensive
Operations/Admin control coverage (Production audit still 61 pending constructed
buttons); guide export/import wire/provenance details; remaining comparison and
presentation evidence; carrier/D5 authority; seven owner capture corrections;
broader Shipping shutdown and its seven real D8-A failures; human/NAS acceptance.
The remaining-acceptance checklist preserves each candidate's scope.

## 8. Immediate next action

Build and run focused packaged-caller D8-A RED for ordinary Auth creation/repair,
exact target, invalid/unreadable sources and preserved authority, while protecting
explicit provisioning and existing credential-validated station transition.

## 9. Critical references

- Architecture v4.11 D8-A, D2 operator sign-in rule8, D18 Event Detail amendment.
- Plan022 current decision/evidence entries; controls1.241.
- `tests/integration/plan022_slice4be_remaining_acceptance.md` and the two result
  records named above.
- `src/Core/Modules/modAuth.bas`: LoadAuth, ResolveAuthWorkbook,
  CloseTransientAuthAfterLoad, ValidateUserCredentialForTarget and
  TryTransitionLegacyS1CapabilitiesForCurrentComputer.
- Explicit LoadAuth callers: modWarehouseBootstrap, Admin modAdminConsole and
  modTesterSetup. Setup can precede current-target selection; do not break it.
- `tests/tooling/Test-Slice4beConfigCommands.ps1`: NewFixture, SelectTarget, Run.
- `tests/tooling/Slice4beAdminUomProbe.ps1`: Invoke-AdminUomResetChoice contains
  the separate native dialog capture helper implicated by the three wrong crops;
  DPI mismatch is a hypothesis, not yet a calibrated diagnosis.
