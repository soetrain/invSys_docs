# 269 B1 Seed profile and bound replay RED

## Goal and release outcome

4be-B B1 must replay one saved built-in Admin Seed How-To in a new Training destination and prove the exact inventory effect of that invocation under D18-REPLAY-01. The profile works. The automatic Seed milestone and operator acceptance remain open.

## Current verified state

Last verified November 25, 2026. The code review branch is `work/d18-picker-save-repro`. The first test is `8ad21694`, Core profile support is `2a427bdc`, and the bound-run RED is `3ff00f5c`. Plan 022 and Controls record the profile result at `880ddcb` on docs branch `work/perplexity-handoff218`, published as `work/d18-picker-save-repro-docs`. The qualification source is catalog 48 at `27cb6d5b` plus the profile registry change. Catalog 49 has separate evidence.

All five packages in `deploy/validation-b1-seed-bound-20261125` compiled. Operations cold-started against dependencies in that directory. The ordinary Excel installation still uses the prior review build. These R1-5 candidate identities distinguish the build tested here.

| XLAM | Build identity | SHA-256 |
| --- | --- | --- |
| Admin | `151cefb048084f1ba65e7598f52a6696` | `3CDDFE9CAA46D28BC3A44710B30751DC2A1D4CE57AA1301C31D55776310CBD9E` |
| Core | `98c42ef2dd174ec9aadbb3bbf1f4f323` | `59629D238FB19EF078D3801735412BCF8B73543044BC6706A9817C0E52538E3D` |
| Designs.Domain | `75562a96efa74a4e8ecdf69c47ff87ae` | `77D2A003F34D3E57A723DF27721F197A80CF438A17917F5BEB46C931E196BBEC` |
| Inventory.Domain | `8c77754ce5d04f31a4eca689885b3c9f` | `BF5234A7E0ECEA6F016CE49ADC43572E5B96D5D5CE2835276155B1C4DA3245D2` |
| Operations | `2809f2fc6a344683a733c0359ebdff9e` | `ACC69BE909592CC0D12B0E146B4C0263DFDB3F95DAB42AE768F2260510AAD4CA` |

## Decisions and constraints

Justin authorized two D13 steps. First establish and fix the unconfigured-profile refusal. Then establish a bound-run RED before changing dispatch or verification. Built-in and uploaded Seed have one recorded control ID. The profile registry accepts the built-in data-set literal `R1_WORKFLOW_KIT`. Source authoring and destination replay use separate generated Training warehouses with ordinary credential validation and capability checks.

## Evidence and traceability

- Final code checkpoint `c5b2afda`. On the complete rebuilt set the focused suite ran **14 PASS / 3 FAIL**, with all three failures named below and no harness exception. Reproduce with `tests/tooling/Test-Slice4beAdminSeedOutcome.ps1 -DeployRoot deploy/validation-b1-seed-bound-20261125 -Phase RED -SeedReplayRun`. The run completed unattended, Excel closed naturally, test settings were restored and all five candidate and tooling pins matched. There were zero Excel Application events 1000/1001/1002 from this full build through final closure. Local controller receipt `8bd4d238db624a4a939a1fdf5b738a94`; typed results `f97dde884a34415099c027e01d5425af`.
- The independent negative case passed. After the positive attempt had ended, ordinary Seed prepared the 21 additional destination groups. The check matched all 21 fresh destination keys to exact event lines and saved entity values, preserved the three bootstrap rows, and compared kit values and catalog fields to the ordinary source result. The earlier event was published with all 21 lines. With the same exact profile, a new Step through run was stopped before step one. Verify read the loaded publication and that new closed journal, returned Failed with REQUIRED_STEP_MISSING, and credited zero terminal sources. The inventory bytes matched before and after Verify.
- Ordinary repeated Seed returned UNCHANGED with zero event references, no new receipt and no duplicate stock. Signing in as a user lacking ADMIN_MAINT or changing the destination after setup prevented Start and preserved inventory. The original source guide and inventory passed their final hash checks. These are distinct from positive replay of Seed or a replayed no-op, which still require the adapter.

- Handoff 267 pins the five onboarding baseline packages. The packaged baseline ran 7 PASS / 1 FAIL at `SeedReplay.Profile.ConfigureOpensForExactSeedGuide`. Configure refused the imported guide and Run reported the missing profile. Source guide and destination inventory were preserved.
- On the profile candidate the imported guide has a new local ID and a transfer origin binding the source record and digest. The saved profile binds its local guide ID, version, record, digest, stable step, adapter ID/version and reviewed literal. Run setup displayed those exact local identities. The immutable run binds that profile and guide, destination, session digest, package builds and recording sequence.
- `SeedReplay.Run.StartNeedsNoReceivingEntity` and `SeedReplay.Run.OrdinaryOwnerAppliedWithFreshEvent` failed. The test selected an existing destination bootstrap entity only to reach the Receiving-only dispatcher. The bound Seed step persisted as Unknown with reason CAPTURE_UNAVAILABLE, an empty activity ID and no source event. The persisted destination inventory matched its baseline. The full proof test also protects `SeedReplay.Run.ExactSavedInventoryAndVerifiedCurrentEvent`.
- The remaining owner chain is the ordinary Demo Inventory Seed handler → `modAdminSeedObservation.Seed` → one paired activity ID and destination Inventory event reference → saved APPLIED receipt, inventory log and entities by exact System_Key. Verify uses SourceEventsApplied and the new recording and published event lines. The test independently reads a snapshot of saved workbook bytes. A new destination already contains three bootstrap groups. The built-in kit should preserve those identities and values and create the other 21 groups with owner-generated keys. Positive replay dispatch, owner outcome and persisted-effect GREEN are all outstanding.
- The product change is confined to `modExecutionInputs.bas`. The focused controller and fixture use `Test-Slice4beAdminSeedOutcome.ps1`, `Test-Slice4beConfigCommands.ps1`, `Slice4beAdminSeedOutcome.ps1`, `Slice4beAdminSeedReplay.ps1` and `Slice4beSeedReplayProof.ps1`. The last file contains the independent saved-state and stopped-run checks.
- Static comparison against handoff 267 yielded 392 components, 6,796 procedures, 147,004 lines (+5), 9 literal and 45 unresolved calls. The additional normalized duplicate candidate is `modExecutionInputs.Supported` and `modShippingActivityCodes.HasInventorySources`. Both select seven control IDs. The actual ID lists differ.

## Do Not Repeat

Rebuilding only Core alongside four copied binaries can pass explicit preload yet fail cold dependency resolution. The intermediate Core hash was `F4B8162A8EBB0B3AB71A747EB8801A1A8640280E8760119CF4F25AFDED3D5B16`. The Windows Application log recorded an Excel VBE7.DLL access violation at 23:50:46 UTC and an Office mismatch report at 23:50:50 during that diagnostic interval. Use the complete rebuilt candidate above. The saved-state fixture must share reads with a writable Excel projection. The evaluation guide reference contains ActionPathId, Version and ContentSha256; the execution profile and run also bind RecordId.

## Assumptions to re-verify

Recheck Excel ownership, loaded paths, five hashes and current desktop before starting another gate. The prior 121/121 Receiving B0 result belongs to the earlier package identities. The changed candidate needs its own B0, Seed, Create/Training, authorization, identity and preservation regressions after focused B1 GREEN.

## Open questions and blockers

Operations references Core and can load without Admin. A fixed primitive Operations-to-Admin Seed call would add one literal Application.Run site (9 to 10, unresolved count unchanged). The requested scope exception covers only that site, with captured Training session, pending step, reviewed kit and live ADMIN_MAINT rechecked at the owner boundary. Prior Shipping approval covered its own fixed bridge.

## Immediate next action

After the bridge decision, make the existing bound-run test pass through the ordinary Seed handler and run its fresh-event, saved-value, refusal and repeat controls.

## Critical references

Architecture v4.11 D18-REPLAY-01 and Plan 022 B1; Controls v1; handoffs 267–268; `modExecutionRun`, `modExecutionBinding`, `frmActionPathRun`, `modAdmin.Seed_DemoInventory`, `modAdminSeedObservation`, and `modEvaluationSources`. The 4be-A picker no-match explanation, full visible onboarding walkthrough, and outstanding operator/recording coverage carry forward. The next increment remains finishing this one Seed adapter and its proof.
