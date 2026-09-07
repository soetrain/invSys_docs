# Slice 4be approved; first activity RED

## 1. Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11 and Plan 022 by
implementing comprehensive Operations/Admin Events and How-To/Diagnostic Action
Paths without regressing accepted workflows, identity or package behavior.

## 2. Current verified state

- Last verified 2026-09-07: code main `5aa8d0f`, pushed; docs main `63c202f`,
  pushed. This handoff/pointer is committed afterward; resolve with Git.
- Architecture D18 is approved and active, including the detailed synthesis.
  Plan 022 is still current. Slice 4be.1 has its first meaningful packaged RED;
  the activity runtime has not been implemented. Deployed XLAMs remain the
  unchanged `b4ce6d9` package set.
- Code tree is clean. Preserve user's unstaged handoff 067 (3 additions /
  3 deletions) and untracked `expert guidance docs/023 Slice 4be Critique.md`.
  Neither was edited or staged. The critique is advisory, not a plan pointer.
- No Excel process remained after focused validation. Recheck before further
  build/deploy. No operational workbook or NAS deployment was changed.

## 3. Decisions and constraints

- User explicitly approved the detailed shared Events/How-To/Diagnostic/Compare
  both contract and asked that critique 023 be considered. Do not ask for the
  same approval again or reinstate curated-versus-captured exclusivity.
- Semantic inheritance is recorded in Architecture D18: lower-level artifacts
  may clarify, discover controls and strengthen tests; contradiction or material
  weakening requires an explicit approved architecture decision before code.
- Adopted critique details: events are facts, named logical owners determine
  effects, stable EventCode differs from RecordId/ActivityId, severity and
  Changed/Unchanged/Unknown effects are structured, text is allowlisted, and
  observation is re-entrancy guarded. Unknown outcome cannot imply rollback.
- Executable Navigate/Retry/Repair/Override Action Paths were not adopted.
  Approved diagnostics evaluate permitted evidence; ordinary workflow owners
  retain all actions. RetryAllowed matches an observed retry, never executes it.
- Preserve D5 Core read/command ownership, D12 headless dependencies, D13,
  immutable exact System_Key, unknown columns, captured workbook binding,
  launcher reuse and all accepted role/Production behavior.

## 4. Evidence and traceability

See [first packaged activity RED](../../invSys_fork/tests/integration/plan022_slice4be_activity_red_results.md).

- Fresh unchanged-package cold-start dependency check and **5/5 compiles PASS**.
- Fresh packaged smoke **81/81 PASS**.
- Focused activity mode: **19 PASS / 12 FAIL**, no harness exception. All 18 D5
  checks pass; direct service call correctly produces no user-control evidence.
- Three real-handler cases lack activity records: Admin Settings changed-value
  save, successful Production UOM Retrieve and denied Retrieve. Each fails four
  evidence assertions (attempt/result, correlation, owner/context/effect,
  redacted payload). This is three missing cases, not 12 separate root causes.
- Tests reuse packaged Admin-generated disposable fixtures and unsaved actual
  form-handler instrumentation. No runtime source was changed to create RED.
  Initial Excel-open precondition failure is excluded from behavioral evidence.
- New tests parse; diffs/whitespace checked. No new live-role/full-chain/layout/
  static/visible UAT completion is claimed. Prior D5 results remain the baseline.

## 5. Do Not Repeat

- Do not label approval pending; approval is recorded in docs `21ff41a`.
- Do not implement only JSON existence to make this small first test pass;
  D18 also requires context/policy checks, hashing/atomic persistence,
  re-entrancy/failure isolation, redaction and comprehensive control coverage.
- Do not overload stable EventCode as an instance ID or use worksheet positions
  to retrofit old Admin history. Existing station-local audit is incomplete.
- Run Excel validators serially. Packaged smoke left one empty process; COM
  inspection verified zero workbooks before it was closed. Never kill an
  operational or unidentified Excel process to make a test start.
- Keep candidate package absolute-reference hazards and explicit compile gates
  in view; a successful build alone is not a compile or restart proof.

## 6. Assumptions to re-verify

Package hashes, Git/Excel state and source-to-control mappings before runtime
work. Physical NAS/station and human UAT remain separate release evidence.

## 7. Open questions and blockers

No contract-approval blocker remains for current D18. Runtime implementation
and remaining 4be.1-4be.6 gates are unfinished. New contradictory discoveries
need their own architecture decision. Overall Release 1 acceptance is open.

## 8. Immediate next action

Implement the approved headless observation boundary and the first catalogued
Admin/Production handlers against the recorded RED, expanding policy/context,
storage-integrity and failure-isolation tests before claiming focused GREEN.

## 9. Critical references

- Architecture `invSys-Design-v4.11.md`, active D18 and Viewer R1 scope.
- Plan 022, current Slice 4be approval/critique disposition, first RED and
  execution table; controls catalog v1.61 and first discovered control entries.
- Code `tests/tooling/Test-Slice4beConfigCommands.ps1 -CheckActivityEvidence`,
  `tests/tooling/Slice4beActivityAssertions.ps1`, linked sanitized evidence.
- `frmAdminSettings.mBtnSaveConfig_Click`,
  `frmProduction.mBtnUomCatalogRetrieve_Click`, `Core.modConfigCommands`.
- Stable ControlIds `ADMIN_SETTINGS_SAVE_VALUE`, `PRODUCTION_UOM_RETRIEVE`;
  logical owner `CORE_CONFIGURATION`; EventCodes listed in the catalog/test.
- Ignored check report `reports/runtime/slice4be-activity/red.json`.
