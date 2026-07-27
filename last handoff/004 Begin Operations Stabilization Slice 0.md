# Begin Operations Stabilization Slice 0

**Created:** 2026-07-26
**Updated:** 2026-07-27

## Goal and release outcome

Begin Slice 0 under the approved normative `System_Key` amendment and converge
on the D12 five-package release with D13 test-first evidence, reproducible
developer tooling, greenfield warehouse generation/seeding, and packaged
Receiving -> Production -> Boxing -> Shipping acceptance proof.

## Current verified state

### Code repository

- Repository: `/mnt/c/Users/justu/source/repos/invSys_fork`
- Branch: `codex/fix-tester-station-nas-setup`
- HEAD: `d93ad7331fbecb7e61e1edbc47056c8b77690f41`
- Latest commit: `d93ad73` — `Add agent guidance and remove LaTeX artifacts`
- Last verified: 2026-07-27
- `AGENTS.md` is committed and contains the new session-start, D13, security,
  completion, and decision-centered handoff instructions.
- `AGENTS.md` is now modified but uncommitted with the matching D14
  `System_Key`, greenfield, extensible-header, and managed-`Condition`
  invariants.
- The old `tikz_playground` `.aux`, `.fdb_latexmk`, `.fls`, `.log`, and `.pdf`
  generated files are deleted in the pushed commit.
- The following tracked files are deleted in the working tree. These deletions
  remain uncommitted and were preserved without staging or reversal:
  - `excel_after_enable.png`
  - `excel_normal_launch.png`
  - `excel_open_xlam.png`
  - `excel_ribbon_check.png`
  - `excel_ribbon_check_max.png`
  - `excel_vbe_error.png`
  - `invSysReceiving_no_punct.png`
  - `receiving_direct_after_fix.png`
  - `receiving_sdk_builder_check.png`
  - `ribbon_test_open.png`
  - `ribbon_test_open_sdk.png`
  - `ribbon_test_open_v2.png`
  - `ribbon_test_open_v3.png`
  - `settings.json`
  - `temp_modAdminConsole_from_xlam.txt`
  - `trusted_path_after_book1.png`
  - `trusted_path_ribbon_check.png`
  - `vbe_foreground.png`
- Windows Git also reports the following modified fixture workbooks, while WSL
  Git did not report them. They were not staged or committed:
  - `tests/fixtures/WH1.invSys.Auth.sample.xlsx`
  - `tests/fixtures/WH1.invSys.Auth.xlsb`
  - `tests/fixtures/WH1.invSys.Config.sample.xlsx`
  - `tests/fixtures/WH1.invSys.Config.xlsb`
  - `tests/fixtures/WH1.invSys.Data.Inventory.xlsb`
  - `tests/fixtures/invSys.Inbox.Production.S1.xlsb`
  - `tests/fixtures/invSys.Inbox.Receiving.S1.xlsb`
  - `tests/fixtures/invSys.Inbox.Shipping.S1.xlsb`

### Documentation repository

- Repository: `/mnt/c/Users/justu/source/repos/invSys_docs`
- Branch: `agent/update-v4-11-operations-packaging`
- HEAD before the current uncommitted documentation edits:
  `d6f9466cb8c0e2899b9abe1e664f918ef3d87ed0`
- Latest commit: `d6f9466` — `Refresh Operations handoff after publish`
- Upstream: `fork/agent/update-v4-11-operations-packaging`
- Pull request: `https://github.com/soetrain/invSys_docs/pull/1`
- Last verified: 2026-07-27
- The previously finalized 020 changes are committed and pushed. They add:
  - modeless main Receiving, Production, and Shipping forms;
  - explicit captured operator-workbook/session binding;
  - duplicate modeless-instance prevention;
  - a non-operational Purchasing stub tab in the main Receiving form;
  - Box Builder and Box Maker tabs in the main Shipping form;
  - removal of separate Box Builder, Box Maker, and Purchasing launch surfaces
    from the Operations ribbon; and
  - packaged navigation/RibbonX gates for those contracts.
- The current-spec, current-plan, and current-handoff pointers are committed.
- `expert guidance docs/021 Overfit and Details.md` is committed as source
  guidance; it is not the designated current implementation plan.
- The current uncommitted 020/handoff edits add the greenfield `System_Key`
  contract, abandon legacy inventory import, and add a dedicated Generate
  Warehouse/demo-seed slice.
- The normative v4.11 specification is modified but uncommitted with D14,
  schema-appendix changes, and Phase 6 acceptance gates.

### Excel state

- Whether Excel is open is unresolved.
- Before any XLAM build, import, deployment, or COM-driven workbook test,
  re-verify Excel state and close the relevant workbooks/add-ins when required.
- Slice 0 schema and synthetic-fixture work should not require an operational
  Excel session.

## Decisions and constraints

- `invSys-Design-v4.11.md` is the normative architecture.
- `020 Operations Stabilization and Developer Tooling Slice Plan.md` is the
  designated current implementation plan.
- D12 requires one deployed `invSys.Operations.xlam` containing internally
  separate Receiving, Production, and Shipping role implementations.
- D13 requires meaningful RED before implementation for Core, Domain, service,
  processor, and packaged form-action contracts.
- The exact managed identity header is `System_Key`. It replaces `ROW`; it is
  generated once, system-wide unique, immutable, and preserved through
  refresh/rebuild boundaries.
- `ITEM_CODE` identifies what an item is. Location, quantity, `Condition`, and
  user-added fields are attributes, not identity.
- Runtime tables define required managed headers while tolerating additional
  end-user columns. `Condition` is a managed inventory header.
- This is a greenfield reset. Do not import, translate, reconcile, or repair old
  business inventory, and do not build a legacy `ROW` mapping.
- Supported test inventory comes from a fresh Admin Generate Warehouse/Create
  Warehouse action and Admin `Seed Demo Inventory`.
- D14 and the matching `AGENTS.md` identity rules satisfy the 020 architectural
  entry gate. Runtime implementation still requires meaningful D13 RED.
- Slice 0 precedes scanner/runtime-extractor implementation and defines
  versioned JSON schemas, synthetic fixtures, redaction tests, deterministic
  JSON-to-Markdown tests, and the dynamic-root registry format.
- Static implementation evidence and runtime-state evidence are separate
  products. Generated evidence describes reality and does not override the
  specification.
- Main Receiving, Production, and Shipping forms are modeless. Actions bind to
  the captured operator workbook/session and never use a newly active workbook
  as implicit authority.
- The main Shipping form uses tabs. Box Builder and Box Maker are tabs, not
  Operations ribbon buttons.
- The main Receiving form includes a visible Purchasing stub tab. It is
  non-operational, performs no writes, introduces no Purchasing contract, and
  has no separate Operations ribbon button or group.
- Handoffs are created only at the end of a chat/session or when explicitly
  requested.

## Evidence and traceability

- No runtime implementation was changed in this session.
- No VBA, packaged XLAM, or Excel integration tests were run.
- The normative specification now contains D14, end-to-end inventory schema
  fields, Phase 6 tasks/tests, and the greenfield Generate Warehouse/seed gate.
- The 020 Markdown was checked with:
  - 12 balanced code fences;
  - no trailing whitespace; and
  - targeted inspection of the `System_Key` entry gate, Slice 0 fixtures,
    greenfield warehouse/seed slice, and renumbered Operations slices.
- `AGENTS.md` was checked with:
  - 4 balanced code fences;
  - no trailing whitespace; and
  - targeted inspection of the matching `System_Key`, extensible-header,
    greenfield, and managed-`Condition` invariants.
- No runtime-state or implementation-manifest report was generated.

## Do Not Repeat

- Do not start Slice 1 or Slice 2 implementation before Slice 0 schemas,
  fixtures, and meaningful failing tests exist.
- Do not begin runtime implementation until the normative specification,
  `AGENTS.md`, 020, and the focused RED tests agree with the approved
  `System_Key` greenfield contract.
- Do not import old inventory, create a legacy `ROW` mapping, or preserve `ROW`
  as a runtime/display compatibility field.
- Do not implement `System_Key` as a rename or copy of an old numeric `ROW`.
- Do not accept Generate Warehouse or `Seed Demo Inventory` output that lacks
  unique keys, `Condition`, extensible-header safety, or projection preservation.
- Do not treat a compile failure, missing fixture, unavailable workbook, or
  broken harness as D13 RED.
- Do not infer the current plan from the highest-numbered guidance file; 021 is
  input material, while the explicit `CURRENT.md` pointer designates 020.
- Do not add Box Builder or Box Maker buttons to the Operations ribbon.
- Do not create a Purchasing ribbon surface or implementation from the
  Receiving placeholder tab.
- Do not implement modeless form actions through `ActiveWorkbook`,
  `ActiveSheet`, or `Selection`.
- Do not discard, restore, stage, or commit the existing dirty-tree changes
  without first resolving their ownership and intended scope.

## Assumptions to Re-verify

- Re-verify that the local branches and HEAD commits above are still current.
- Re-verify that the three committed pointer files resolve correctly in the
  next environment.
- Re-verify Excel open/closed state before any Excel automation.
- Documentation PR #1 was open and draft when last verified on 2026-07-26.
- The D14/`AGENTS.md` amendment is present in the working trees but remains
  uncommitted and must be reviewed as one contract.
- The proposed tool entry points in 020 remain guidance until Slice 0 freezes
  their schemas and fixtures.

## Open questions and blockers

- No architectural question blocks Slice 0 after review of the D14/`AGENTS.md`
  amendment.
- The normative specification, `AGENTS.md`, 020, and this handoff contain
  coordinated uncommitted changes that need an intentional commit/push.
- The remaining code-repository deletions and Windows-Git-only fixture
  modifications need ownership confirmation before any future commit.

## Immediate next action

Review the coordinated D14/`AGENTS.md`/020 amendment, then write and run the
first meaningful failing Slice 0 static-maintenance schema and deterministic
JSON-to-Markdown fixture test before implementing the scanner.

## Critical references

- `../invSys_fork/AGENTS.md`
- `0 plan docs/xlam_invSys/CURRENT_SPEC.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`
- `expert guidance docs/CURRENT.md`
- `expert guidance docs/020 Operations Stabilization and Developer Tooling Slice Plan.md`
- `expert guidance docs/021 Overfit and Details.md`
- `last handoff/CURRENT.md`
- v4.11 D12 — Operations Packaging Consolidation
- v4.11 D13 — Test-First Development for Core, Domain, Service, and Form-Action
  Contracts
- v4.11 D14 — System-Wide Entity Identity and Extensible Headers
- 020 Entry gate — normative `System_Key` decision
- 020 Slice 0 — Tool contracts, schemas, and synthetic fixtures
- 020 Slice 4 — Greenfield warehouse generation and demo inventory seeding
- `../invSys_fork/src/Core/Modules/modWarehouseBootstrap.bas`
- `../invSys_fork/src/Admin/Modules/modAdmin.bas` — `Seed_DemoInventory`
- `../invSys_fork/src/Admin/Forms/frmSeedInventory.frm`
- `../invSys_fork/src/Core/Modules/modRoleEventWriter.bas` — current
  `QueueMigrationSeedEvent` dependency to remove from greenfield runtime seeding
- Proposed Tool A entry point: `tools/inventory-vba-surface.ps1`
- Proposed Tool B entry point: `tools/export-invsys-runtime-state.ps1`
