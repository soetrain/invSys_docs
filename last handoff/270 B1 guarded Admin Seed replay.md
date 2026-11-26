# 270 B1 guarded Admin Seed replay

## Goal and release outcome

4be-B B1 covers replay of the built-in Admin Seed How-To into a separate Training warehouse. D18-REPLAY-01 requires evidence of that invocation's effect in the saved destination.

## Current verified state

Last verified November 25, 2026. Code `b2865aefe89c0fc9648834221228cc7f23f3d9d9` on `work/d18-picker-save-repro` adds the bridge and focused tests to the profile and bound-run RED checkpoint `c5b2afda`. The qualification build uses handoff 269's isolated catalog 48 lineage, with its profile change and all six changed product files. All five R1-5 XLAMs were rebuilt together in `deploy/validation-b1-seed-bridge-01`. Operations initialized cold and all five pristine projects compiled. These identities identify the candidate used for the runs.

| XLAM | Build identity | SHA-256 |
| --- | --- | --- |
| Admin | `63b7b7846f894fd7acef45ee00629ca4` | `D9A7DD57E9DDF76556FE9860ADD67953350404844FE236DC2B64E28D21FB2A05` |
| Core | `6613640618d445ba970c7a7057dbc834` | `50126214113788BF1850D949A5A3EFD5DC63A9D705F8130398AA47D0B6E2ACFA` |
| Designs.Domain | `b4e0e83b5d5c4aacb322a15e63154860` | `140D96BA938D25FAF0EE8067588E7845015C33F702635C298C3FDC5683D8E955` |
| Inventory.Domain | `75d677c146c04fb7b8435a41e83c163f` | `8C33665F30A4C0FF46C7BD987C27084FD9071D7E9672EE5FBF8C556D9EA7B71A` |
| Operations | `439cdc79bad5422297e3d3a7972ebbde` | `755FC3C2D8908F964A0455DAAC5F28CBBB73F84623D36F18FA23F3D104FF66ED` |

Plan 022 and Controls were updated at docs commit `f4bbc694eea58991c3b3a4dc288598383f5ac11b` on `work/perplexity-handoff218`, published as `work/d18-picker-save-repro-docs`.

## Decisions and constraints

Start derives item prerequisites from the configured steps. Seed uses the reviewed built-in kit. Receiving still requires a valid local inventory entity. The fixed call targets `'invSys.Admin.xlam'!modAdminSeedReplay.ApplyBuiltIn` with three strings (setup token, stable pending step ID, kit ID). Core checks the active Training session and destination, loaded package identities, saved run and recording, guide/profile and pending step. Admin rechecks the selected actor, target, runtime and ADMIN_MAINT, takes a single claim and invokes the ordinary form's built-in selection and observed Seed operation. Missing or incompatible Admin fails binding. A lost or uncertain response receives no retry. The owner may return APPLIED or UNCHANGED. Verify independently evaluates attributable evidence.

## Evidence and traceability

- Handoff 269 recorded **14 pass / 3 fail** on the complete bound-profile candidate. All three now pass through the packaged Run How-To handler: `SeedReplay.Run.StartNeedsNoReceivingEntity`, `SeedReplay.Run.OrdinaryOwnerAppliedWithFreshEvent` and `SeedReplay.Run.ExactSavedInventoryAndVerifiedCurrentEvent`.
- The first bridge run passed 20/20. Its final cleanup attempted a test-only Settings helper on a freshly reopened, uninstrumented Admin copy. The fixture now closes Settings before unloading Admin. The permanent closed-recording/build-identity assertion also passed on the rerun: **21/21**, controller `bac861b0cda847e9b5f3f54d22fa048e`, worker `be1a5dbe7b404f32858ae48c0bb0f54c`. Reproduce with `tests/tooling/Test-Slice4beAdminSeedOutcome.ps1 -DeployRoot deploy/validation-b1-seed-bridge-01 -Phase GREEN -SeedReplayRun`.
- The positive check reads the destination's closed recording, published event and saved inventory independently. The run observation and owner pair share an activity ID and reference the same submitted Inventory event. The saved APPLIED receipt, 21 exact event lines and associated System_Key values agree. The three bootstrap groups retain their exact identities and values. Each of the other 21 groups has the expected values and a new destination key. Verify credits this event to the current run. Source guide and source inventory pins match.
- A second run stopped before Seed with matching stock and all 21 earlier lines published returned REQUIRED_STEP_MISSING, credited zero terminal sources and left inventory bytes intact. Ordinary and replayed repeat Seed returned UNCHANGED, with zero fresh event references, new receipts or duplicate stock. The replayed no-op earned zero terminal sources. Altered kit, pending-step ID, setup token and duplicate bridge requests were refused. Loss of ADMIN_MAINT or a destination change after setup prevented Start and writes. Unloading Admin prevented Seed setup.
- Static scan against handoff 269: 393 components (+1), 6,799 procedures (+3), 147,102 lines (+98), 194 duplicate-body candidates (unchanged). There are 10 literal Application.Run sites (+1) and 45 unresolved (unchanged). The extra site is the fixed Operations-to-Admin Seed call. The Admin entry point is a cross-package root.

| Suite on the candidate above | Passed | Failed |
| --- | ---: | ---: |
| Focused B1 built-in Seed replay | 21 | 0 |
| Receiving B0, including run prerequisites and busy close | 121 | 0 |
| Ordinary Admin Seed, recording and workflow boundaries | 181 | 0 |
| Create Warehouse owner | 110 | 0 |
| Authorization, read/write and user administration | 259 | 0 |
| Initial Admin and Training onboarding | 59 | 0 |
| Ordered full chain, System_Key uniqueness, exact balances, event replay and restart reconciliation | 33 | 0 |
| **Total** | **784** | **0** |

The focused controller and all six retained gates ended with natural Excel closure and restored settings. No modal was dismissed. The focused controller verified its complete tooling and package pins. The final comparison matched all nine changed source/test pins and all 15 XLAM pins across the previous candidate, bridge candidate and installed review packages. The Windows Application log recorded zero Excel error, hang or reporting entries (1000/1001/1002) from 00:46:38 to 01:32:46 UTC on November 26 (November 25 local). No Excel process remained. The retained gate receipts are under `b0`, `seed`, `create`, `authority`, `training` and `chain` in the local qualification directory.

## Do Not Repeat

Use a coherent build. Handoff 269 recorded a VBE7.DLL access violation at 23:50:46 UTC and a Windows Error Reporting mismatch at 23:50:50 UTC during the earlier mixed-package diagnostic. Its cause remains undetermined.

## Assumptions to re-verify

Recheck the current five package identities, Excel ownership, Training context and test settings before another runtime gate. Catalog 49 and the installed reader-reopen build have separate qualification histories.

## Open questions and blockers

Visible operator acceptance of the complete workflow, the 4be-A picker no-match explanation, and outstanding recording coverage remain open. This increment covers one B1 adapter. The broader B1 work in Plan 022 includes nested entry, interruption and partial outcomes. Release 1 acceptance and promotion remain owner decisions. Code PR 17 and docs PR 8 remain closed. Their historical bases are `work/d14-packaged-chain-coverage` and `work/excel-shutdown-timing` respectively. The feature branches carry the commits.

## Immediate next action

Exercise the built-in Seed How-To from Run through Verify with an operator on a separate Training destination and record visible acceptance in Controls.

## Critical references

Architecture v4.11 D18-REPLAY-01, Plan 022 B1, Controls v1, handoff 269, `frmActionPathRun.DispatchSteps`, `modExecutionBinding`, `modExecutionRun.PendingSeed`, `modAdminSeedReplay.ApplyBuiltIn`, `modAdminSeedObservation.Seed`, `modEvaluationSources`, `Slice4beAdminSeedReplay.ps1` and `Slice4beSeedReplayProof.ps1`. The local qualification directory is `reports/runtime/b1-seed-bridge-qualification`.
