# Goal and release outcome

Continue invSys Release 1 acceptance under Architecture v4.11. Slice 4au fixes
the Receiving runtime error 380 and header overlap while preserving the visibly
accepted Slice 4at Production location-stock allocation.

# Current verified state

- Code `main`: `59b5122` (`Fix Receiving result list column limit`).
- Docs `main`: `6e41e52` (`Record Receiving result list correction`) before this
  handoff commit.
- Active slice: Plan 022 Slice 4au, automated GREEN; visible Receiving
  confirmation pending.
- Last verified 2026-08-30: Excel is closed. Code retains the user's modified
  `AGENTS.md`; docs retain the user's modified Architecture v4.11 note and
  untracked `invSys dev notes.txt`.

# Decisions and constraints

- MSForms ListBox `.List` is limited to ten safe columns. Receive Item Results
  therefore shows exactly Code, Item, UOM, Available, Location, blank Capacity,
  Lot, Condition, Description, and Vendor.
- The representative exact `System_Key` is not displayed; it is held in a
  row-aligned collection and consumed by the real Receive/Return action.
- Capacity remains an inert Release 1 stub. No capacity schema or behavior was
  introduced.
- All Receiving header labels are single-line/non-wrapping.
- Production's one-row-per-item/location stock projection is visibly accepted
  and must not be reverted to receipt-event rows.

# Evidence and traceability

- Focused Slice 4au RED `0/7`, GREEN `7/7`.
- Packaged Receiving durability/search/header action `1/1`, including
  `SearchRowsLoaded=True`, `HiddenSystemKeyMap=True`,
  `TenColumnItemResults=True`, and `HeadersSingleLine=True`.
- Receiving stabilization `10/10`; prior Slice 4at `7/7`; packaged XLAM
  `81/81`; Ribbon/VBA compile `142/142`; live roles `47/47`; Release 1 chain
  `30/30`; deterministic static `19/19`; reviewed growth `13/13`.
- Static metrics: 154 components, 5,224 procedures, 1,050 candidates.
- Current NAS rerun was not available; prior verified NAS evidence is `16/16`.

# Do Not Repeat

- Do not add an eleventh ListBox column for hidden identity; it causes runtime
  error 380 on non-empty results.
- The first live-role run had a generic Excel COM exception; a clean-process
  rerun passed `47/47`, so no code change was made for that transient failure.

# Assumptions to Re-verify

- Confirm the installed `deploy/current` packages are the new build before
  visible testing.
- Confirm the user's real warehouse search contains rows; the packaged fixture
  uses a synthetic non-empty row to exercise the exact fill routine.

# Open questions and blockers

- Visible confirmation remains open for Receiving search results and readable
  one-line Received Tally/Aggregate Received headers.
- Broader Production multi-process visible acceptance remains ongoing.

# Immediate next action

Install/reuse the current packaged add-ins, open Receiving, search a managed
item, and confirm results load without error 380 and no header overlaps the
lists.

# Critical references

- `src/Receiving/Forms/frmReceiving.frm`
- `tests/tooling/Test-Plan022Slice4auReceivingListbox.ps1`
- `tools/validate_plan022_packaged_launchers.ps1`
- Plan 022 Slice 4au
- `invSys-Controls-v1.md` v1.37
