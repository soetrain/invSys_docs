# D8-INV-01 tightened; approval pending

## Goal and release outcome

Complete Plan0224be-A recording/guidance and early Receiving B0 while preserving
Release1. This turn applies the user's three D8-INV-01 proposal tightenings;
the proposal remains unapproved and comprehensive4be-A remains incomplete.

## Current verified state

Last verified2026-10-15 Pacific, repository/proposal review only; no runtime
test, Excel inspection, build or deployment was performed in this revision turn.

- Code `codex/aggregation-authority`, `a4d14ba4`: Inventory receipt now contains
  the written unresolved Config diagnosis for trial115.
- Docs `codex/tighten-d8-agg-01`, `98f63d7` before this handoff/pointer commit:
  revised proposed D8-INV-01, Plan022 and Controls1.636 agree.
- Product source remains1c9bff87; diagnostic baseline55868837; unpromoted
  `validation-aggregation-recording-02` packages unchanged. Last verified runtime
  state and scoped technical qualifications are in handoff187/code receipts;
  recheck Excel closure before any build. No runtime freshness is inferred here.
- At handoff preparation only this handoff/CURRENT need committing. Push both
  branches to their configured `fork` remotes and verify clean status before exit.

## Decisions and constraints

- **D8-INV-01 remains PROPOSED; approval pending.** The user requested tightening,
  not implementation or approval. Preserve the user's return-on-new-proposal rule.
- Tightening1: remove the instruction to align D14's earlier positive-only phrase.
  Instead D13 must prove zero-quantity modal Add and worksheet ADD remain accepted
  with an active entity and exact saved quantity. D14's entire locked section is
  byte-for-byte unchanged. This supersedes handoff187's alignment-on-approval text.
- Tightening2: modal Add/Edit/Delete jointly revalidate captured catalog source,
  selection and affected entity set after the same yields/before writes. External
  change/deletion/replacement refuses; same-Item-Code replacement cannot silently
  rebind. Add also checks its catalog source before a new entity exists. Fresh
  context/reopening is required; returning to the old selection cannot revive it.
- Tightening3: each RED cohort requires a written `Config diagnosis` section in
  the Inventory receipt, bound to its exact product/test commits and controller/
  worker IDs. Record evidence plus established cause or explicit unresolved status,
  remaining questions and next diagnostic. Focused GREEN cannot close an unresolved
  finding; different-source comparisons must be labeled separately.
- All earlier headless authorization, D8 TTL, D14 identity, preserved workflows,
  honest partial retry and desktop security constraints remain. No new ControlId,
  replay adapter, receipt store, quantity allocation or migration is authorized.

## Evidence and traceability

- Trial115 remains35 PASS/10 FAIL with45 retained checks, no harness exception;
  exact signed-out persisted create/ROW/host findings and unresolved Config change
  retain source1c9bff87/test55868837 scope. No new behavioral RED/GREEN is claimed.
- `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`, section
  `Config diagnosis`, explicitly records cause **unresolved** for controller
  `f997acde371142d79f956ca8c5502b25`, worker `ce2b57b7540e490fb57431fc1775792f`.
  It distinguishes observed Config hash change/Auth preservation from an unproved
  semantic cause and specifies a redacted, read-only phase comparison next.
- Diff checks,134 local Markdown target checks, three-request/provenance checks
  and exact D14 preservation pass. Specification changes are confined to the
  pending D8-INV-01 section. No runtime implementation changed; proposal review
  does not manufacture a new D13 RED or promote the existing45-check diagnostic
  into coverage of the expanded matrix.

## Do Not Repeat

Do not restore the implicit D14 editorial authorization, assume same-SKU catalog
replacement is safe, or close Config when focused GREEN passes. Do not infer a
cause from file hashes or reuse a different-source comparison as the RED diagnosis.
Keep handoff187's runtime failure/oracle/closure constraints and historical scopes.

## Assumptions to Re-verify

Read pointers, actual user approval and Git state. Before runtime work, recheck
Excel/owned-job closure and passive desktop access. Newly required modal replacement
and zero-quantity cases are specified, not already demonstrated by trial115.

## Open questions and blockers

D8-INV-01 approval remains required. Config cause remains unresolved. Expanded
actual-action/yield/phase RED, implementation, recording/guidance, Release1 fixture
conformance and remaining coverage remain open; prior technical gates retain the
exact scopes in handoff187 and the maintained acceptance index.

## Immediate next action

Obtain approval of the tightened D8-INV-01, then expand actual modal/worksheet RED for replacement refusal and zero-quantity preservation while carrying the cohort-bound Config diagnosis into D13 review before product edits.

## Critical references

- Current v4.11 D8-INV-01 (line5745 at this revision), D8/D13/D14; Plan0224be-A;
  Controls1.636; handoff187 for qualified runtime gates and exact receipt locations.
- Code `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`,
  especially `Config diagnosis`; `plan022_slice4be_remaining_acceptance.md`.
- `tests/tooling/Slice4beAdminInventoryEntry.ps1` and
  `Test-Slice4beConfigCommands.ps1 -AdminInventoryEntryOnly`;
  `modAdmin.Add_InventoryItem`, `frmAddInventoryItem`, `modAdminInventoryWorksheet`.
