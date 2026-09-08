# Receiving navigation and identity GREEN

## 1. Goal and release outcome

Achieve Release 1 user acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths.
The full Goal remains active and incomplete. This session completes the 4be.1
Receiving navigation/selection technical checkpoint and a discovered D14 identity
blocker; it does not complete comprehensive 4be.1 or human UAT.

## 2. Current verified state

- Last verified 2026-09-08: code main **77a0851**, docs main **bda1456**, both
  pushed. This handoff/pointer is committed afterward.
- Final candidate: `deploy/validation-receiving-navigation-identity`; **771/771**
  full packaged activity/form checks and all technical release gates GREEN.
- All five final candidate hashes, the first navigation candidate hashes and
  accepted catalog-5 lifecycle hashes remain unchanged. Accepted catalog-4
  freshness hashes also reverified unchanged. Accepted deployment/current and
  NAS were not replaced.
- Excel is closed; recheck before builds/tests. Final Production interval has
  zero Excel Application Error events.
- Code tree is clean. Preserve unrelated docs: handoff 067 has 3 additions/3
  deletions; critique 023 remains untracked. Neither was edited or staged.

## 3. Decisions and constraints

D18 semantic inheritance remains approved; no new approval blocker. Architecture,
Plan 022 and controls **v1.78** are synchronized. Navigation refinement was pushed
as docs **c4cbde6** before runtime implementation. The identity correction restores
existing D14; it changes no saved identity, migration rule or Domain rejection.

Catalog 6 adds thirteen Receiving page/list/choice identities, Navigation class,
RECEIVING_NAVIGATION owner, RECEIVE_POST eligibility, collection default off.
Explicit current policy can enable collection; capture flag alone cannot enable
it outside recording. Actual mouse/keyboard selection records REQUESTED and
UI-only SELECTED/Unchanged, fixed captions and empty source references. Input
values, programmatic selection, initialization, internal refresh/detail and
window mechanics are excluded. Stale/closed captured context rejects before
detail-owner entry. Optional tracking failure stays visible without retrying the
owner. Purchasing remains the stub; recording integration remains pending.

Core's new modSystemIdentity extracts the existing Windows GUID implementation.
Role entity/event creation and archive collision suffix use it; Training retains
its representation/error wording. Do not restore the per-call Randomize fallback.
All D5/D12/D13/D14, headless authority, exact System_Key, unknown-column and
captured-workbook/launcher reuse rules remain binding.

## 4. Evidence and traceability

[Navigation/identity evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_navigation_results.md)
contains exact candidate hashes, RED/GREEN, gate commands and preserved failures.

- Navigation: first full RED 632 PASS / 74 FAIL retains all 596 previous GREEN;
  expanded mouse/keyboard focused RED 118/122; first candidate 253 then 257
  focused and 767 full GREEN. Final corrected candidate is **771/771**, every
  earlier check identity retained, no duplicate check names.
- Discovered blocker: chain 29/30, Admin Seed rejected DUPLICATE_SYSTEM_KEY ->
  Core fallback burst 35,017 unique / 14,983 duplicates in 50,000 calls -> D14 ->
  `Test-PackagedSystemKeyGeneration.ps1`. After the fix, both 50,000-key cases
  have zero duplicates/blanks, package unchanged, **3/3 GREEN**; chain **30/30**.
- Five builds/explicit VBE compiles/cold starts PASS; smoke **86/86**, live roles
  **48/48**, Create Warehouse source integration **15/15**, Viewer PASS,
  three requested sizes/five pages/native transitions PASS, launchers **3/3**,
  full ProductionReusable plus clean restart **2/2**. No reduced Production flags.
- Receiving and Production captures inspected. This is agent evidence, not human
  comparison or physical multi-station/NAS acceptance.
- Static: 175 components, 5,500 procedures, 1,097 candidates, 195 duplicate groups,
  45 unresolved calls/eight literal targets; all 28 prior size caps hold.
  Receiving form 1,200 lines; modTS_Received 1,575; modRoleEventWriter 3,064.
  Explicit reviewed exception retains 17 native roots and three short normalized
  event-adapter groups; no duplicated business logic or dynamic-call growth.
- Relevant source/tooling checks pass; exact counts are in the linked evidence.

## 5. Do Not Repeat

- Do not count compiler/harness/input-delivery failure as D13 RED. Unsaved VBA
  seams cannot use reserved Tab; TabStrip mouse events require ByVal Index.
- New shared source dependencies must enter every explicit source-harness import
  list. All sixteen lists now include modSystemIdentity. The first corrected
  chain paused in a missing-dependency compiler dialog; exact test ownership was
  verified, only the chain parent stopped, and VBE Reset allowed normal cleanup.
  No Excel process was killed in that recovery; the clean rerun is the gate.
- Wait for a validator to finish before reading its result. After launcher GREEN,
  the next guard briefly saw Excel exiting; no second validator started. Wait
  for actual process exit, never kill unidentified Excel.
- Identical minimum/default images appeared blank when displayed together. Check
  file hashes and view individually before treating this as bad capture/runtime.
  The minimum clamps to default geometry; originals and recapture are valid.
- Preserve lifecycle safety from handoff 078: native Hide plus synchronous finish
  and flag-only invalidation; no nested Unload or releasing the cached form inside
  QueryClose. Keep process-owned native modal observation; never accept OK/Cancel.
- Do not relocate compiled XLAMs with absolute dependencies, retry away failed
  evidence, or substitute diagnostic-only/reduced Production runs for full GREEN.

## 6. Assumptions to re-verify

Git/Excel state, exact candidate hashes, actual control reachability and captured
ownership. Older lifecycle ntdll failures remain causally unproven; this final
candidate's complete Production run passed without a new crash.

## 7. Open questions and blockers

No current approval/environment blocker. Receiving worksheet Confirm reachability
and launcher denial observation remain pending, followed by other Operations/Admin
controls and shared publication. Event Tracking Settings/profiles/preferences,
comprehensive Viewer, recording/conclusions, guide management and both comparison
presentations remain incomplete. Physical multi-station/NAS/user UAT remains.

## 8. Immediate next action

Add packaged reachability evidence for ReceivedTally's btnConfirmWrites under the
actual Operations launcher before deciding its explicit control identity or
documented exclusion and implementing the next D13 coverage increment.

## 9. Critical references

- Architecture v4.11 D18/D14; Plan 022 4be.1-4be.6; controls v1.78; linked evidence.
- `frmReceiving.NavigationSelection`, `modReceivingNavigation`,
  `cReceivingSelectionInput`, `modReceivingNavigationCodes`, `modActivityCatalog`.
- `modSystemIdentity.NewId`, `modRoleEventWriter.CreateSystemKey`,
  `modTrainingWire.NewId`; `Test-PackagedSystemKeyGeneration.ps1`.
- `Test-Slice4beConfigCommands.ps1` with every activity/foundation/Receiving/
  staging/local/lifecycle/navigation switch; `Slice4beReceivingNavigation.ps1`.
- Ignored `reports/runtime/slice4be-system-key/activity-results.json`,
  `full-chain-results.md`, `package-hashes.json`,
  `production/production-reusable-production.md`, `layout-recapture/`.
- Next surface: `modTS_Received.InitializeReceivingUiForWorkbook`,
  `EnforceReceivingSupportSheetsHidden`, `EnsureReceivingButtons`, `ConfirmWrites`.
  Support sheet is VeryHidden when another visible sheet exists; the compatibility
  entry remains public. Do not infer user clicks from arbitrary macro calls.
