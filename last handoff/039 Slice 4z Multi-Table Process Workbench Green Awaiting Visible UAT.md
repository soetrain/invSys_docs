# Goal and release outcome

Achieve Release 1 acceptance under Architecture v4.11; Plan 022 Slice 4z now
provides the automated-GREEN multi-table Process worksheet workbench while
preserving the accepted Viewer and completed Seed, Receiving, Shipping, Boxing,
and packaged Operations behavior.

# Current verified state

- Code `main` is pushed at `75627ec` (`Complete Slice 4z Process worksheet
  workbench`). Documentation `main` is pushed at `7ac72fd` (`Reconcile Slice 4z
  Process workbench contract`). Last verified: 2026-08-24.
- Active Slice 4z automated gates are complete; operator-visible workbench UAT
  remains open. Excel is closed.
- `deploy/current` contains the exact tested rebuilt Core and Operations XLAMs.
- Preserved unrelated changes remain unstaged: code `AGENTS.md` and the user's
  NAS/server note at the top of `invSys-Design-v4.11.md`.

# Decisions and constraints

- The supplied screenshot is positive Viewer acceptance: the uploaded/seeded
  sample is visible. Do not reopen a Viewer defect from that image.
- Process, Recipe, Requirement, and Output IDs remain locked generated
  three-character Base-36 values. Output Design ID/version are generated from
  Process/Output identity; operator-authored output Item Code is hidden/removed.
- Process Designer has separate **Create Process Table** and **Retrieve Selected
  Process** actions. Any number of tables may coexist; retrieval uses selection
  in the captured workbook and deletes only the successfully imported table.
- INPUT Percent and Basis Qty are managed formulas. ALTERNATIVE rows expose
  Requirement ID, acceptable managed item, and hidden managed SKU, with the
  existing Core Production item picker.
- Core/Domain remain headless authority; worksheet tables are CSV-friendly
  staging surfaces only. Production Run - Tree remains experimental.

# Evidence and traceability

- RED: `tests/integration/plan022_slice4z_process_workbench_red_results.md`.
  GREEN: `tests/integration/plan022_slice4z_process_workbench_green_results.md`.
- Focused source 7/7; packaged public launcher and clean restart 2/2. Three
  tables coexisted; selected retrieval left two; both survived restart and were
  then retrieved one at a time.
- Formula proof: 611.2 lb basis; 16.4%, 32.7%, 1.8%, 49.1%; total 100.0%.
- Regressions: XLAM 74/74; Ribbon/compile 142/142; live roles 47/47; ordered R1
  chain 30/30; launcher contracts 24/24; dedicated NAS 16/16; deterministic
  static baseline 19/19; reviewed cleanup 13/13.
- Layout RED found an expanded-page Retrieve/Description overlap. The widened,
  relocated Retrieve control is GREEN across all five pages at minimum,
  default, expanded, minimize/restore, and maximize sizes.
- Static metrics: 153 components, 5,041 procedures, 1,036 candidates, 189
  duplicate groups, 8 literal `Application.Run` targets, 45 unresolved dynamic
  calls.

# Do Not Repeat

- Do not restore the Slice 4y Edit/Retrieve toggle or singleton-table state.
- Do not treat the accepted Viewer screenshot as evidence that the item is
  missing.
- Two full-chain attempts ended 28/29 only because the read-only extractor
  attached to a stale background Excel process. Closing that exact process made
  the extractor succeed; the clean rerun was 30/30. No Product fix was needed.

# Assumptions to Re-verify

- The user has not yet visibly accepted multi-table creation, CSV paste/reformat,
  acceptable-item search, selected retrieval, or Recipe Designer ordering on
  the real saved Production workbook.
- `deploy/current` is tested, but installation/deployment into the user's normal
  Excel add-in location must be confirmed before visible UAT.

# Open questions and blockers

- No automated blocker. Remaining Slice 4z gate is operator-visible acceptance;
  broader Plan 022 visible saved-workbook/NAS Production acceptance remains open.

# Immediate next action

Deploy/reuse the tested packages, then run visible Production UAT by creating at
least two Process tables, pasting the 611.2 lb formulation, using acceptable-item
search, retrieving only the selected table, and ordering the retrieved Processes
in Recipe Designer.

# Critical references

- `src/Production/Forms/frmProduction.frm`
- `src/Production/Modules/modProductionProcessWorksheet.bas`
- `src/Production/ClassModules/cProductionAppEvents.cls`
- `src/Core/ClassModules/cDynItemSearch.cls`
- `tools/validate_plan022_packaged_launchers.ps1`
- `tests/tooling/Test-Plan022Slice4zProcessWorkbench.ps1`
- Architecture v4.11 D15; Plan 022 Slice 4z; controls catalog section 8.1
