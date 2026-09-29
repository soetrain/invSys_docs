# Desktop stop and approved UOM reuse continuation

## Goal and release outcome

Complete Release 1 acceptance under Architecture v4.11 and Plan 022, including
Slice 4be comprehensive Event Viewer/Action Path coverage. The goal remains
incomplete and was explicitly **paused on 2026-09-29** after desktop error 5,
as the user instructed. Do not resume tests or implementation without user resume.

## Current verified state

Last verified: **2026-09-29**. Code `main` **d6cc64b**, docs `main` **509c243**,
both pushed before this handoff. Code is clean; preserve the unrelated modified
`last handoff/067 Partial Goal Action Path NAS Contract.md` and untracked
`expert guidance docs/023 Slice 4be Critique.md`. No Excel process remains.
The owned finite desktop observer was stopped after recording the failure.

Current frozen candidate: `deploy/validation-production-instructions-typed`,
runtime source **4f2d52b**. Five packages and settings were preserved during
subsequent diagnostics. Controls version **1.261**, Production coverage **1.6**.
No runtime fix for the new UOM defect has been implemented. No rebuild/deployment
with relevant Excel files open; reverify package pins before dependent tests.

## Decisions and constraints

- **Latest user approval, 2026-09-29:** approved the documented UOM reuse proposal:
  Edit UOM Catalog on Sheet reopens existing staging without resetting edits or
  custom columns; load saved catalog only for a new empty workbench. Existing
  drafts do not automatically reload later saved-catalog changes. The proposal
  also specifies reopening the identifiable region after successful Retrieve
  unlists it, normalized managed headers, and rejection of ambiguous shapes
  before mutation. **Architecture, Plan and controls still label the proposal
  pending. Record this approval in those three artifacts before implementation.**
- This approval arrived after the desktop stop; it does not resume the goal.
- User will attend screen-dependent work, but attendance did not prevent the
  observed failure. On another desktop error 5, stop/pause immediately. Do not
  infer a lock cause, change policies, or treat full access as desktop access.
- Preserve the user's disabled `invSys.StationUpdate`; no reboot is authorized.
- D5/D13/D18, semantic inheritance, captured-workbook binding, immutable exact
  `System_Key`, unknown columns, headless authority and packaged launcher reuse
  remain binding. Diagnostic passes do not establish release acceptance.

## Evidence and traceability

Current candidate, September 29:

| Gate | Result and limits |
| --- | --- |
| Five instruction controls | Focused RED 105 PASS/306 FAIL became GREEN 411/411; paired Action Paths 105/105. Five instrumented compiles, normal closure and preservation. Six principal images reviewed. |
| Release 1 chain | 32/32; live roles 48/48; Create Warehouse 15/15; normal closure and preservation. |
| Regressions | Settings 202/202; draft/path 390/390; lifecycle 615/615; native cancellation 94/94, all normal closure/preservation. |
| Settings observations | 467/467, but assisted cleanup; normal shutdown remains open. |
| Smoke | 86/86; normal exit unproven because validator can terminate Excel without a receipt. |
| Layout/static | Three complete images reviewed after DPI capture fix; geometry/actions pass. 255 components, 6065 procedures, 133334 lines, 191 duplicate groups; 28 caps unchanged, 9 literal/45 unresolved dynamic calls. Three schemas and 305 PowerShell parses pass. |
| Cold reusable Production | Unmodified RunOnly fails at batch contract call: RPC 0x800706BE, Excel ntdll c0000028. Prior 67 observations not reestablished; full/restart open. |
| UOM staging | Actual packaged Send handler: **49 PASS/5 behavioral FAIL**. Repeat Edit deletes both custom columns, custom values/formulas/order and unrelated worksheet content. Normal closure, five instrumented compiles and preservation. |

UOM root cause: `modProductionUomCatalog.SendUomCatalogToWorksheet` unlists the
existing table then calls `ws.Cells.Clear`. Governing requirement: existing
header-extension preservation rule plus approved reuse decision. Protecting test:
`tests/tooling/Test-Slice4beProductionUomStaging.ps1`, underlying actual-handler
probe `Slice4beProductionUomStaging.ps1`. RED report:
`reports/runtime/slice4be-production-uom-staging/1423544f53da42d68aa4afcca4e2f1a1/red.json`.
The earlier locked-workbook hash failure was harness setup, not behavioral RED.
Retrieval currently sends all columns by ordinal; the planned normalized-header
retrieval test has **not been added or run**.

Desktop stop: first failing sample **2026-09-29 19:45:13 UTC** (12:45:13 Pacific)
reports CursorError=5, InputDesktopError=0, CaptureError=6. At detection there
were 33 affected samples. This proves cursor access failure, not its cause or an
input-desktop error 5. Private evidence:
`reports/runtime/rdp-desktop-probe-2a6a5fac7fc84325ba65754536f21e74/samples.jsonl`.
No new Excel test had started when the failure was detected.

## Do Not Repeat

- Do not rerun completed broad gates merely to recover context.
- Saved compilation into a separate candidate did not fix cold failure; that
  candidate crashed earlier at opening Production. Do not promote it.
- Source tracing, VBE-only preparation and native debugger attachment can yield
  passing assertions while requiring forced cleanup. None establishes a fix.
- Bounded COM-reference release and post-report error-reference clearing did not
  establish normal packaged shutdown; no diagnostic cleanup variant was adopted
  into the ordinary validator. See batch-boundary evidence before further work.
- Raised VBA/native exception code 5 is distinct from Windows desktop error 5.
- Do not expose raw report details, credentials, fixture values or machine logs.

## Assumptions to re-verify

User resume and desktop availability; no Excel open; frozen candidate pins at
`reports/runtime/production-instruction-controller/babcce28be534358ba247e93d1c044df/package-pins.json`;
both repositories' branch/status. Desktop observer is stopped, not monitoring.

## Open questions and blockers

UOM preservation is RED; retrieval/header/reopen tests and approved implementation
remain. Production coverage is 18/68 constructed buttons registered, 50 pending;
14 unconstructed legacy buttons and 34 non-button handlers remain accounted for.
Other Operations/Admin coverage and human/NAS acceptance remain open. Cold native
crash and unassisted shutdown are unresolved. The earlier 6.5/10 estimate is not
a measured acceptance percentage.

## Immediate next action

After explicit resume, record the received UOM approval in Architecture, Plan and
controls, then expand the actual-handler RED for normalized retrieval and approved
draft reopening before changing runtime implementation.

## Critical references

- Authority pointers: `0 plan docs/xlam_invSys/CURRENT_SPEC.md`,
  `expert guidance docs/CURRENT.md`; controls `0 plan docs/xlam_invSys/invSys-Controls-v1.md`.
- Pending clause to mark approved: `invSys-Design-v4.11.md`, UOM Catalog section
  around line 3550; matching Plan 022 and controls paragraphs.
- Code evidence: `tests/integration/plan022_slice4be_remaining_acceptance.md`,
  `plan022_slice4be_production_instruction_results.md`,
  `plan022_slice4be_production_uom_staging_results.md`,
  `plan022_slice4be_production_batch_boundary_results.md` in the same directory.
- Runtime: `src/Production/Modules/modProductionUomCatalogWorksheet.bas`
  (VBA module name `modProductionUomCatalog`), `frmProduction` actual
  `mBtnUomCatalogSend_Click` / Retrieve handlers; headless Core publication remains
  authoritative and its authorization must not be inferred from staging permission.
