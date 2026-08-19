# Goal and release outcome

Complete Release 1 operator workflows against the dedicated NAS warehouse; this
Slice 4q follow-up stabilizes the live **Add Disposition** transaction and makes
any remaining Excel write failure actionable.

# Current verified state

- Code: `main` at `1ed25d9` (`Stabilize Receiving disposition staging`), pushed
  2026-08-19.
- Docs: `main` at `2a1e506` before this handoff commit.
- Active slice: Plan 022 Slice 4q, automated GREEN; visible `WHT7025AE`
  Return/Dump retest remains.
- Rebuilt and committed `invSys.Core.xlam` and `invSys.Operations.xlam` plus the
  deployment manifest.
- Excel is fully closed. Last verified 2026-08-19.
- Preserved unstaged user changes: code `AGENTS.md`; the NAS/server note at the
  top of `invSys-Design-v4.11.md`.

# Decisions and constraints

- RETURN/DUMP continues to deplete existing exact `System_Key` entities; no
  replacement identity or inbound-return behavior was introduced.
- Programmatic disposition allocation and aggregate rebuild run with Excel
  events suppressed and restore the caller's prior event state.
- Manual Received Tally quantity synchronization uses the worksheet's captured
  workbook, not `ActiveWorkbook`, and suppresses its own recursive write event.
- Disposition group scans stop at the first mismatched attribute so unrelated
  inventory rows cannot disrupt a valid exact-group allocation.
- A staging failure reports stage, error number, sanitized source, and
  description; it no longer collapses to generic error 1004 text.

# Evidence and traceability

- Symptom: **Add Disposition** for 5 of 50 available DAMAGED Packing Tape Strip
  reported `Application-defined or object-defined error` and staged no row.
- Read-only diagnosis: captured workbook, selected key, quantity, exact group,
  zero-row staging table, saved full inventory, and complete live inventory all
  staged successfully in disposable copies; no operational workbook was
  mutated or saved.
- Same-handler protected-write test: RED returned generic text; GREEN reports
  exact stage/number/source while the ordinary action still stages the existing
  identity.
- Focused Receiving range 290-296: 7/7; Slice 4o/4p/4q static contracts: 5/5,
  6/6, 6/6; packaged Returns: 5/5.
- Packaged XLAM: 74/74; live role workflow: 46/46; ordered Release 1 chain:
  30/30.
- Static maintenance: Slice 3 19/19 and Slice 12 11/11; 150 components, 4,701
  procedures, 965 candidates, 8 literal `Application.Run` targets, 47 unresolved
  expressions, and 184 duplicate candidates.

# Do Not Repeat

- Do not use `Application.ActiveWorkbook` from a modeless form or worksheet
  event when the source workbook is already known.
- Do not allow programmatic staging writes to re-enter Received Tally change
  handling.
- Do not diagnose this as a missing/duplicate `System_Key`; the exact live key
  and group passed isolated allocation.

# Assumptions to Re-verify

- Excel will load the rebuilt add-ins from `deploy/current` on next start.
- The live error was caused by session-only write/event interaction; the new
  event-isolated transaction still requires the operator's exact visible retest.

# Open questions and blockers

- Visible UAT has not yet confirmed that DAMAGED Packing Tape Strip stages and
  confirms through RETURN or DUMP in `WHT7025AE`.
- If staging still fails, record the new complete status text; it now identifies
  the exact failing stage and source.

# Immediate next action

Start Excel, repeat Disposition Ref `Dis104` (or a new reference) for 5 DAMAGED
Packing Tape Strips, confirm **Return Tally** receives the row, then confirm the
disposition and verify Returns plus Inventory Viewer decrease from 50 to 45.

# Critical references

- `src/Receiving/Modules/modTS_Received.bas:StageInventoryDispositionForWorkbook`
- `src/Receiving/Modules/modTS_Received.bas:StageDispositionAllocation`
- `src/Receiving/Modules/modTS_Received.bas:HandleReceivingSheetChange`
- `src/Receiving/Forms/frmReceiving.frm:TestStageProtectedDispositionActionForWorkbook`
- `tests/unit/TestReceivingStabilization.bas:TestReceivingReturns_StagesExistingDispositionIdentityThroughFormAction`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, section 7.1
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4q
