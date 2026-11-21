# Handoff 221: Retained receipts and CREATE regressions

Date: 2026-11-05 Pacific. Predecessor: handoff 220.

## Goal and release outcome

Complete invSys Release 1 Slice 4be-A under Architecture v4.11 and Plan 022.
This increment advances approved D8-INV-EDIT-01/02, D8-INV-CONFIG-01 and D13;
full 4be-A and Release 1 acceptance remain open.

## Current verified state

Last verified: 2026-11-05 5:36 PM Pacific. Code worktree
`invSys_edit_receipts` is clean on `work/d8-inv-edit-receipts` at
`1e1524b5c01e6e2a81e1a509586ed0b099e949d8`, pushed to
`soetrain/invSys_fork:work/d8-inv-edit-receipts`. The main code checkout has
the user's unrelated modified `AGENTS.md`. Before this handoff commit, docs
branch `work/perplexity-handoff218` is at
`538f9d59cbee7e1acde084f8c81b12ef9a9dadbc`, pushed to
`soetrain/invSys_docs`; its pre-existing deletion/edit and untracked
guidance files remain untouched. This handoff and pointer are the only intended
uncommitted docs changes before publication. Excel and test controllers are
closed; the desktop monitor was stopped normally. No XLAMs were rebuilt or
promoted in this increment.

## Decisions and constraints

No new architecture or control contract was introduced. Candidate
`deploy/validation-inventory-config-03` remains pinned. The test-only owned
Seed Inventory fixture close, actual captured Upload handler, typed Core receipt
check, private Inventory receipt, and exact `System_Key`/event identity rules
remain as in handoff 220. Controls 1.731 and Plan 022 record this evidence.
No WPR/ETW capture, security change or assisted Excel shutdown occurred.

## Evidence and traceability

- Packaged `RetainedReceipts`: **81/81 GREEN**, controller
  `green-d5c80885c6c84984900e1f463fcb8120`, worker
  `f4eda0ba4a3241809d53b23eea95fdab`. Catalog-only, no-op quantity,
  quantity change, and failed queue/reopened retry passed; exact saved
  command/source, pre-queue planned event ID and applied/log proof were checked.
- Packaged EDIT creation `Success`: **88/88 GREEN**, controller
  `green-fc0444a923084b5a844aea183d06d86e`, worker
  `178e220df41d41ed9256907668b603e7`. Fresh creation, reopened retry and
  six identity-loss refusals passed without replacement identities.
- Packaged EDIT creation `QueueFailure`: **34/34 GREEN**, controller
  `green-f8fda55c6c1645cb99b72222ec7d7058`, worker
  `afe9d79a354a4a47a1a7db9c023f219d`. Failed queue created no entity or
  imported status; reopened retry kept command/event/key and applied once.
- Every run compiled five instrumented packages, closed Excel unassisted,
  restored settings, preserved all five package hashes, and had zero delayed
  Excel Application 1000/1001/1002 faults. Fresh cursor proof succeeded; the
  stopped monitor `b335dacd045c4d8f915ab09d5cae2e6f` reported zero error-5,
  desktop-lock and mouse-input failures. Config/Auth, presentation, custom
  columns and read-only inspection checks passed in the named runs.
- The QueueFailure worker spent about **3 minutes 44 seconds** between
  `AfterCollectBeforeFinalizers` and `AfterGarbageCollection`. This preserves
  evidence of intermittent harness finalization delay. Exact shutdown records
  and generated reports are ignored and remain local. Code audit:
  `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`.
  Handoff 220 retains focused CREATE 140/140, Config 123/123 and missing-default
  54/54 evidence and static baseline; no product source changed this turn.

## Do Not Repeat

The earlier combined QuantityPhase run required assisted teardown and is not
qualified GREEN; retain its isolated 140-check matrix. The earlier 20/1
RetainedReceipts desktop-stop result remains excluded. Do not treat the Office
update or this run's unassisted closure as a universal shutdown fix. Do not
retry the rejected docs PR-creation action through another route: automatic
approval review returned **“blocked by policy”** without a specific reason.

## Assumptions to Re-verify

Before more Excel runtime work, recheck Excel closure, desktop unlock/cursor
movement, candidate hashes and current worktrees. Code draft PR4 targets
`soetrain/invSys_fork:main` with soetrain head. Docs draft PR5 targets
`soetrain/invSys_docs:work/excel-shutdown-timing` but still has the older
`justinwj` head `54cd913`; the new docs branch is published on soetrain and
does not update PR5. Recheck remote PR heads when publication resumes.

## Open questions and blockers

The docs PR route remains unresolved after the policy rejection. Broader
Inventory mutation/regression, live-role, visible operator, full Release 1
chain and final static/acceptance gates remain open. No new specification
approval is pending.

## Immediate next action

After a fresh desktop/cursor/pin check, run the existing packaged
`run-inventory-edit-creation.ps1 -Phase GREEN -Case ApplyFailure` on candidate
03 as the next isolated D13 regression, then qualify its natural closure.

## Critical references

Architecture v4.11 D8-INV-EDIT-01/02 and D8-INV-CONFIG-01; Plan 022; Controls
1.731; `tests/tooling/Slice4beInventoryEditCreation.ps1`,
`tests/tooling/Slice4beInventoryEditReceipts.ps1`,
`reports/runtime/run-inventory-edit-creation.ps1`,
`deploy/validation-inventory-config-03`.
