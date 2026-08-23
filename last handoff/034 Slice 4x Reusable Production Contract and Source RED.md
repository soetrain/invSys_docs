# Goal and release outcome

Complete Release 1 acceptance by implementing Architecture v4.11 D15 reusable
Production Processes and Recipe graphs without regressing the packaged
Operations launchers or completed Seed, Receiving, Shipping, Boxing, and Viewer
behavior.

# Current verified state

- Code branch `main`, commit `009bd20` (`Record reusable Production design RED`), pushed to GitHub.
- Documentation branch `main`, commit `f641eb5` (`Define reusable Production process architecture`), pushed to GitHub.
- Active slice: Plan 022 Slice 4x -- reusable Production Processes and Recipe graphs.
- Contract reconciliation is complete in the normative spec, current plan, and controls catalog; no Production VBA/schema implementation change has begun.
- Focused source RED is recorded at 1/10. Only the preserved public launcher/captured-workbook boundary passes.
- Excel is fully closed. Last verified: 2026-08-23.
- The user's pre-existing code `AGENTS.md` change and NAS/server note near the top of `invSys-Design-v4.11.md` remain uncommitted and preserved.

# Decisions and constraints

- Process and Recipe are distinct immutable, versioned Designs Domain definitions.
- Process Designer replaces part of the former Recipe Builder contract and owns requirements, acceptable SKU alternatives, instructions, one or more outputs, and save/release/obsolete/reuse.
- Recipe Designer selects exact released Process versions and validates output-to-requirement connections, quantities, execution order, unresolved inputs, compatibility, and cycles.
- Production Run - List allocates exact current inventory `System_Key` entities, creates each executed output under a distinct new key, consumes routed intermediate output by that same key, and leaves unconnected balances as finished/co-product inventory.
- The approved form has five pages: Process Designer, Recipe Designer, Ingredients Assignment, Production Run - List, and experimental out-of-scope Production Run - Tree.
- Existing launcher reuse, captured saved-workbook binding, `0.001%`-`1000%` scaling, persistence summaries, headless Core/Domain authority, and current GREEN role/Viewer behavior remain regression locks.

# Evidence and traceability

- Documentation validation: D15 once, Slice 4x once, five-page target present, balanced Markdown fences, and `git diff --check` clean.
- Source RED command: `& .\tests\tooling\Test-Plan022Slice4xReusableProduction.ps1`
- Source RED result: 1 passed / 9 failed in `tests/integration/plan022_slice4x_reusable_production_red_results.md`.
- Expected failures cover the five-page form, Process lifecycle/minimum output, Recipe graph validation, Process requirement alternatives, typed multi-output run/completion, Designs projections/events, and Viewer Production labels.

# Do Not Repeat

- Do not relabel the existing Recipe Builder controls as Process/Recipe designers without implementing separate Designs Domain authority and handlers.
- Do not extend the singular `OutputSystemKey` session as a compatibility shortcut; D15 requires one new key per executed output.
- Do not treat an output definition, SKU, Recipe, or Process identity as physical inventory identity.
- Do not begin implementation before packaged public form-action and Designs Domain RED are recorded.

# Assumptions to Re-verify

- The first source RED is intentionally supplemental; the next gate must execute the deployed public launcher and actual form handlers.
- Exact new runtime control IDs may change after handler-test design, but operator wording and D15 behaviors are fixed.
- Legacy recipe import remains explicit design-definition conversion only; no silent runtime fallback is allowed when Designs are enabled.

# Open questions and blockers

- No external blocker. Packaged public-handler RED and Designs Domain RED remain to be created before implementation.

# Immediate next action

Create and run a packaged Slice 4x harness that enters through `mProduction.BtnOpenProductionForm` and the actual Process/Recipe/ingredient/run handlers, recording meaningful RED against the current four-page form before editing Production or Designs implementation.

# Critical references

- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`: D15
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`: Slice 4x
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`: Production sections 8.1-8.5
- `tests/tooling/Test-Plan022Slice4xReusableProduction.ps1`
- `tests/integration/plan022_slice4x_reusable_production_red_results.md`
- `src/Production/Modules/mProduction.bas`: `BtnOpenProductionForm`
- `src/Production/Forms/frmProduction.frm`
- `src/Production/ClassModules/cProductionRunSession.cls`
- `src/Production/Modules/modProductionCompletionService.bas`
- `src/DesignsDomain/Modules/modDesignsApply.bas`
- `src/DesignsDomain/Modules/modDesignsSchema.bas`
