# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4af makes Recipe identity ready before operator actions and keeps the automated
Recipe version editable.

# Current verified state

- Last verified 2026-08-26: code `main` at `767ccf5`; normative/plan/control
  docs at `53864e0`; both are ready to push before this handoff commit.
- Active slice: Plan 022 Slice 4af, automated GREEN; visible Recipe Designer
  retest is pending.
- Rebuilt five-package set is in `deploy/current`; Operations and Admin startup
  add-ins are registered; Excel is closed.
- Preserved uncommitted user work: code `AGENTS.md` and the NAS/server note at
  the top of Architecture v4.11.

# Decisions and constraints

- A blank Recipe draft receives the next collision-checked three-character
  Base-36 Recipe ID and proposed version `1` after form initialization, New
  Recipe, and Clear.
- Save Draft, Validate, and Release defensively initialize either missing value
  before Recipe graph validation.
- Recipe ID remains locked. Recipe Version is editable and accepts a positive
  whole number; a valid nonblank operator value is not replaced.
- Existing saved ID/version pairs remain immutable and Domain-authoritative;
  no lifecycle action overwrites an unavailable version.

# Evidence and traceability

- Focused source RED `0/6`; focused GREEN `6/6`.
- Packaged Production callback/restart `2/2` records
  `RecipeIdentityInitialized=True`, `RecipeIdGenerated=True`,
  `RecipeVersionGenerated=True`, `RecipeIdLocked=True`,
  `RecipeVersionEditable=True`, and `EditedRecipeVersionRetained=True` through
  the actual New Recipe, Save Draft, and Release handlers.
- Regressions: XLAM `74/74`, Ribbon/compile `142/142`, live roles `47/47`,
  ordered Release 1 chain `30/30`, launcher contracts `24/24`, dedicated NAS
  `16/16`, deterministic static `19/19`, reviewed growth `13/13`.
- Static metrics: 153 components, 5,081 procedures, 1,038 candidates; literal
  and unresolved dynamic-call ratchets did not regress.

# Do Not Repeat

- Do not relock Recipe Version; only Recipe ID is operator-inaccessible.
- Do not require New Recipe before identity appears; initialization is part of
  the public packaged contract.
- Do not overwrite a valid operator-entered version while applying the blank-
  identity fallback.
- `Test-Plan022Slice4xReusableProduction.ps1` is the historical RED recorder,
  not the current GREEN regression; use ProductionReusable packaged evidence.

# Assumptions to Re-verify

- Confirm normal Excel startup loads the registered `deploy/current` package
  hashes rather than a cached older add-in instance.
- The next visible Recipe ID depends on the saved Recipe namespace and may be
  greater than `001`; only the three-character Base-36 form is fixed.

# Open questions and blockers

- No automated blocker remains. Visible confirmation of prepopulated Recipe ID
  and editable Version is unresolved.

# Immediate next action

Open Production Recipe Designer and confirm Recipe ID and Version are already
populated, change Version if desired, then Save Draft and Release the staged
Recipe graph.

# Critical references

- `src/Production/Forms/frmProduction.frm:EnsureRecipeDraftIdentity`
- `src/Production/Forms/frmProduction.frm:ValidateRecipeDraft`
- `src/Production/Forms/frmProduction.frm:ExerciseReusableProductionFormActions`
- `tests/tooling/Test-Plan022Slice4afRecipeIdentity.ps1`
- `tests/integration/plan022_slice4af_recipe_identity_green_results.md`
- `tools/validate_plan022_packaged_launchers.ps1:ProductionReusable`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md` Slice 4af
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` Slice 4af and section 8.2
