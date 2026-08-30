# Slice 4at Location Stock Green Awaiting Visible Acceptance

## Goal and release outcome

Advance Release 1 Production acceptance by showing managed stock per warehouse
location while preserving exact receipt keys for consumption, and stub the
future Receiving capacity display under Architecture v4.11 D14/D15.

## Current verified state

- Code `main`: implementation commit `0ac0243` (`Aggregate Production stock by
  location`).
- Docs `main`: contract/evidence commit `02c3fd2` (`Define location stock
  allocation contract`); this handoff follows it.
- Active slice: Plan 022 Slice 4at, automated local/isolated-package GREEN;
  visible acceptance and the current NAS rerun remain open.
- `deploy/current` is rebuilt and Excel startup registration points to its
  Operations/Admin add-ins. Excel is closed. Last verified 2026-08-30.
- Preserved uncommitted user work: code `AGENTS.md`; docs Architecture v4.11
  NAS/server note; untracked `invSys dev notes.txt`.

## Decisions and constraints

- Production Run shows one stock row per managed
  SKU/UOM/Location/Condition; receipt `System_Key` values are hidden there.
- The selected stock row retains a hidden representative key only for lookup.
  **Apply** atomically and deterministically expands the requested quantity
  across sufficient exact keys before Check In.
- Inventory Check, stale/overdraw/location validation, completion events, and
  inventory mutation remain exact-key authoritative.
- Ingredients Assignment keeps exact identity visible with a 190-point key
  column; saved alternatives remain SKU/item-code definitions.
- Receiving **Capacity (coming later)** is blank and inert: no validation,
  persistence, configuration, or inventory effect.

## Evidence and traceability

- Focused Slice 4at: RED `0/7`, GREEN `7/7`.
- Packaged Production public actions/run/restart: `2/2`, including
  `LocationStockBuckets=True` and `LocationStockExactExpansion=True` against
  seven authoritative receipt keys.
- Packaged Receiving durability: `1/1`, including `CapacityStub=True`.
- Adjacent Production source: `6/6`, `7/7`, `8/8`, `6/6`; Receiving `10/10`;
  workflow readiness `18/18`.
- Launcher contracts `24/24`; packaged XLAM `81/81`; Ribbon/compile `142/142`;
  live roles `47/47`; ordered Release 1 `30/30`; static `19/19`; reviewed
  growth `13/13`.
- Static metrics: 154 components, 5,223 procedures, 1,050 scanner candidates.
- Evidence: `tests/integration/plan022_slice4at_location_stock_green_results.md`.
- Packaged reports:
  `reports/runtime/plan022-slice4at-production-green2/production-reusable-production.md`
  and
  `reports/runtime/plan022-slice4at-receiving/receiving-launcher-durability.md`.
- Current dedicated-NAS rerun did not start because the configured test root
  was unavailable. Prior NAS `16/16` is the last verified result and is not
  claimed as current Slice 4at evidence.

## Do Not Repeat

- Do not display one Production allocation row per Receiving event or invent
  an aggregate `System_Key`.
- Do not replace exact-key event payloads with SKU/location keys.
- Do not treat Capacity as functional until units, scope, persistence, and
  enforcement receive a later approved contract.
- A first packaged run hung because a Boolean variable had the same
  case-insensitive name as `AssignmentSystemKeyReadable`; the variable is now
  `assignmentKeyReadable`.
- A projection-only multi-key test fixture was rebuilt from the authoritative
  ledger into one key. The final fixture writes its split entities to
  `tblInventoryLog`; use the final GREEN2 report.

## Assumptions to Re-verify

- The configured NAS test root becomes reachable before claiming current
  `16/16` evidence.
- The user's Cassia Oil receipt entities share SKU, UOM, CLEARVIEW, and
  Condition; differing Condition values intentionally produce separate rows.

## Open questions and blockers

- Visible operator confirmation is unresolved.
- Current dedicated-NAS evidence is blocked by the unavailable configured test
  root; no product defect was observed.

## Immediate next action

Open Excel normally, launch Production, load the multi-Process Recipe, and
confirm Cassia Oil appears once at CLEARVIEW with summed Available stock; then
open Ingredients Assignment and Receiving to confirm the readable key and
blank Capacity column.

## Critical references

- `src/Production/Modules/modProductionReusableRun.bas`:
  `ReusableRunPaletteRows`, `ApplyReusableRunStockAllocation`.
- `src/Production/Forms/frmProduction.frm`:
  `ApplySelectedRunPaletteSplit`, `BuildAssignmentPage`, `BuildLoaderPage`.
- `src/Receiving/Forms/frmReceiving.frm`:
  `FillReceiveItemResults`, `ApplyReceivingHeaderLayout`.
- `tests/tooling/Test-Plan022Slice4atLocationStock.ps1`.
- `tools/validate_plan022_packaged_launchers.ps1`.
- Architecture v4.11 D14/D15, Plan 022 Slice 4at, controls v1 Slice 4at.
