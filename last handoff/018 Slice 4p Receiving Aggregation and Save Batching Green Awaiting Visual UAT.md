# Goal and release outcome

Complete Release 1 operator workflows against the dedicated NAS warehouse; Slice 4p restores Receiving/Returns aggregation semantics, corrects Return surfaces, and bounds workbook persistence during sign-in and multi-line confirmation.

# Current verified state

- Code: `main` at `0721689` (`Batch receiving aggregates and persistence`), pushed 2026-08-19.
- Docs: `main` at `0a492c8` (`Document receiving aggregation and save batching`), pushed 2026-08-19 before this handoff commit.
- Active slice: Plan 022 Slice 4p, automated GREEN; visible `WHT7025AE` UAT remains.
- Rebuilt and committed `invSys.Core.xlam`, `invSys.Inventory.Domain.xlam`, and `invSys.Operations.xlam` plus the deployment manifest.
- Excel is fully closed. Last verified 2026-08-19.
- Preserved unstaged user changes: code `AGENTS.md`; the NAS/server note at the top of `invSys-Design-v4.11.md`.

# Decisions and constraints

- `ReceivedTally` is the separately keyed posting authority. `AggregateReceived` is a read-only display projection and never supplies queue/log identity.
- Aggregate rows group by receipt type, item code, UOM, location, lot, and Condition. They sum quantity and concatenate distinct references/reasons in first-seen order; different Conditions remain separate.
- Returns uses the `RECEIVE` event boundary and exact titles **Return Entries History**, **Return Tally**, and **Aggregate Returns**.
- Condition is visible in both Receiving and Return item results and remains established in Receiving, not Inventory Viewer.
- Healthy Config/Auth reads do not normalize or save valid workbooks. Missing schema still uses the repair/save path.

# Evidence and traceability

- Behavior locks 289-294 GREEN 6/6 after RED for the four new contracts.
- Focused Slice 4p static contract 6/6; packaged Slice 4o/4p public form actions 5/5.
- Packaged XLAM 74/74; RibbonX 142/142; live role workflow 46/46; ordered standalone Release 1 chain 30/30.
- Three-event Receiving processor proof reports exactly three persistence saves: canonical inventory, one batched outbox append, and one inbox-status save.
- Static Slice 3 19/19 and Slice 12 11/11; 150 components, 4,686 procedures, and 964 scanner candidates.

# Do Not Repeat

- Do not restore reference or return reason to the aggregate grouping key; those values concatenate for display.
- Do not submit aggregate rows or derive `System_Key`/`EventId` from them.
- Do not suppress real Config/Auth schema repair merely to avoid save notifications.
- One combined live-workflow/full-chain invocation encountered a temporary-path harness exception after an orphaned validation Excel process. Stop only that orphan and run the ordered chain standalone; the clean rerun passed 30/30.

# Assumptions to Re-verify

- The installed add-ins Excel loads are the rebuilt `deploy/current` package set.
- The dedicated NAS connection may make individual bounded saves perceptible even though save count is no longer row-proportional.

# Open questions and blockers

- Visible UAT has not yet confirmed cross-reference aggregation, Condition separation, Return labels/Condition, or bounded Saving notifications against `WHT7025AE`.
- The complete Release 1 user-side workflow checkpoint remains open after this focused Receiving retest.

# Immediate next action

In `WHT7025AE`, stage the same item under two PO/BOL references with one shared Condition and one different Condition, confirm the aggregate totals/reference concatenation, then repeat on Returns while counting Saving notices.

# Critical references

- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`, Receiving aggregate and persistence rules
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4p
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, section 7
- `src/Receiving/Forms/frmReceiving.frm`
- `src/Receiving/Modules/modTS_Received.bas`
- `src/Receiving/Modules/modReceivingPostingService.bas`
- `src/Core/Modules/modRoleEventWriter.bas`
- `src/Core/Modules/modProcessor.bas`
- `src/Core/Modules/modRuntimeWorkbooks.bas`
- `tests/tooling/Test-Plan022Slice4pReceiving.ps1`
- `tools/validate_plan022_slice4o_packaged.ps1`
