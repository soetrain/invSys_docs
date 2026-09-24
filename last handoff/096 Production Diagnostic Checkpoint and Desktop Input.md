# Slice 4be: Production diagnostics and desktop input

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11 and Plan022 with comprehensive
Operations/Admin events and both Action Path presentations. **Goal active and
incomplete.** This checkpoint proves the six Production draft actions through
recording, publication, Viewer selection and explicit diagnostic expectations.

## 2. Current verified state

Last verified: 2026-09-24 05:13:47 UTC. Both repositories main; code **9cf03db**,
docs **e0f4e75** before this handoff, both pushed. Controls **1.238**.
Candidate: `deploy/validation-production-diagnostics`; accepted deployment untouched.
Excel closed, all workers terminal, settings restored. Five package hashes pinned.

Only modExpectationDraft and modEvaluationMatches change among 244 compiled
components: offer the catalog-supported VALIDATED choice and explicitly map the
four Production NEW/CLEAR controls to STAGED and two VALIDATE controls to VALIDATED.
Catalog12/68 IDs and runtime observation contracts remain unchanged.

Unrelated edits preserved: frmEventDetail.frm +18 measurement lines, Locked=True;
docs handoff067 +3/-3; untracked critique023. Existing Event Detail bytes remain
in candidate packages but are not committed. No other task edits remain after
this checkpoint/handoff is committed.

## 3. Decisions and constraints

Normative D18, Plan022 and controls were updated before runtime edits. These exact
owner mappings inherit the approved Settings/local-command distinction; no new
authority, permission or architecture exception. Command completion never asserts
Domain application. Empty source references cannot satisfy SourceEventsApplied.
Journal order is exact. Existing publication wire preserves complete original
records/timestamps and group ordering, not filesystem-fed detail-line positions.

Pending, not approved/effective: D8-A Auth read/provisioning separation, Event Detail
Locked=False for non-editable selection/scrolling, D5 carrier authority resolution,
and exact guide-transfer wire/provenance details. Do not infer approval from goal
continuation. User authorizes closing Excel; preserve unsaved operational work.

## 4. Evidence and traceability

Primary evidence: code `tests/integration/plan022_slice4be_production_paths_results.md`.

- Real released-Process Recipe fixture: RED67 PASS/177 expected FAIL -> GREEN244/244,
  retaining preceding234, on unchanged captions candidate.
- Path diagnostic RED370 PASS/eight expected FAIL -> GREEN390/390. All378 RED and
 244 prior GREEN identities retained; actual public form/editor/Evaluate handlers.
- GREEN05:03:35--05:13:30 UTC; unassisted eventual closure, settings restored,
  five package/215 test hashes and three unrelated edits preserved. Zero Excel
  Application1000/1001/1002 events through05:13:47 UTC.
- Five build/compile/cold-start gates; only two Core components changed. Static:
  251 components/6050 procedures/133025 lines (+4);9literal/45unresolved dynamic
  calls,193duplicate groups and28oversized ratchets unchanged. Three schemas valid,
  274 PowerShell parses; source layout7/7+8/8.
- First visible attempt315 PASS/one cursor harness exception, not GREEN. Two
  reviewed Production designer captures show valid Process/Recipe drafts. No
  Event Detail or diagnostic images accepted. Normal unassisted closure.

Private receipt prefix `reports/runtime/production-diagnostics-`: focused-verification,
red-verification, component-verification, static-verification, package-pins,
focused-test-pins, unrelated-pins. GREEN controller `434daf393a934932acddba2d84af2304`;
report `slice4be-production-designer/5e01a29b8a7e42ed8b41322696ae639c`.
Visible attempt controller `ec9b2cf4f1f34b21bef8a913630c9042`; images/report
`slice4be-production-designer/6fc9621d823b4290a1c33840e9a726ad`.

## 5. Do Not Repeat

Adapter calibrations are not product RED:278PASS/fixture exception needed the real
Recipe Refresh before recording;338PASS/sixFAIL needed missing-choice cancellation
before selecting a nonexistent terminal step. The second needed explicit VBA
dialog termination and assisted cleanup. Preserve both attempts; do not treat
their totals as completed gates. Do not compare published line insertion order
to journal order. Do not retry visible capture without changed desktop access.
Do not rerun unchanged successful behavior merely because time elapsed.

## 6. Assumptions to Re-verify

At05:03:34 UTC a non-elevated shell reproduces GetCursorPos error5 on WinSta0/Default.
At05:04:25 WTS reports active RDP, but thread UOI_IO=False. The opened input-desktop
handle also reports False at05:04:54. At05:06:55 the ordinary process successfully
opens WinSta0 with READATTRIBUTES|WRITEATTRIBUTES(0x12). At05:12:20 cursor error5 and
UOI_IO=False persist. These facts support an input-desktop problem; they do not
prove a specific RDP transition or that elevation fixes it. No permissions/UAC/
session configuration changed. Earlier standard-privilege Settings and two current
Production captures succeeded. Recheck current desktop/session and Excel state.

## 7. Open questions and blockers

User was asked whether the PC is unlocked at console, connected by RDP or locked/
disconnected; no new answer received. User already said they use console and RDP.
Ask them to keep the desktop unlocked and RDP open/unminimized. See Microsoft links
in the evidence record. Remaining visible, owner-command and full-chain/live-role
gates need stable desktop input; full-chain confirmation helpers send keystrokes.
Previous460owner/202Settings/67reusable and32chain/48live/15CreateWarehouse results
retain earlier candidate scope. Broad coverage/comparison/transfer/human acceptance
remain open in `plan022_slice4be_remaining_acceptance.md`.

## 8. Immediate next action

Read-only verify restored cursor/input-desktop access, then run the current packaged
Production path gate with `-CheckPaths -CapturePaths` to complete the missing visible
evidence before the prepared owner and full-chain regressions.

## 9. Critical references

- `tests/tooling/Test-Slice4beProductionDesigner.ps1` and `Slice4beProductionPaths.ps1`.
- `src/Core/Modules/modExpectationDraft.bas`, `modEvaluationMatches.bas`.
- Private prepared `reports/runtime/run-production-diagnostics-gate.ps1 -Gate owner`
  or `-Gate chain`; uses settings restoration and frozen package pins. Never run
  two Excel gates together. No rebuild is required for unchanged code.
- Current authority pointers and controls catalog remain the governing documents.
