# 267 Initial Admin and scoped guide-author onboarding

## Goal and release outcome

Justin approved the November 24 proposal at docs `52184a4` and subsequently authorized cleanup of proven attempt-owned files and empty directories after an initial Admin write fails. The increment implements ordinary authorized Create → destination Sign In → scoped Grant → Author/Edit/Transfer → Revoke under D8-CREATE-01, D8-USERS-01 and D18-GUIDE-INPUT-01. Visible operator acceptance and release decisions remain open.

## Current verified state

- Code branch `work/d18-picker-save-repro`. Tests `b67b03e5`, product `e90c3bbb`, supplemental tests `5d6cd860`, stale-choice feedback `d2085dfb`, direct Create fixture `cc0602ba`. Docs branch `work/perplexity-handoff218` publishes to `work/d18-picker-save-repro-docs`. Architecture, Plan 022 and Controls record approval and recovery at `b7d33bc` and `a4c63d8`. Controls and Plan record implementation at `35ec517`.
- Gate packages came from isolated catalog 48 source `27cb6d5b66985b9dd106a590a038a468e83fbb90`, equivalent to applying only the two onboarding product commits to `e47bc394`. Candidate directory is `deploy/validation-onboarding-v4-20261124` in the feature worktree. Five pristine projects compiled and Operations cold-started. Catalog 49 has separate source and evidence. Feature branches are the review route. The owner-closed code PR 17 and docs PR 8 remain closed.
- Seven runtime files changed. `src/Admin/Forms/frmCreateWarehouse.frm` owns the two masked fields and clears transient input. `src/Core/Modules/modWarehouseCreateCommands.bas` validates both entries and binds a single bootstrap call to live source authority. `src/Core/Modules/modWarehouseBootstrap.bas` stages the initial account before publication artifacts. `src/Core/Modules/modAuth.bas` shares the existing credential representation/writer and refreshes persisted authority before protected guide actions. `src/Admin/Forms/frmCreateDeleteUser.frm` owns exact-scope display and explicit guide intent. `src/Admin/Modules/modAdminUsersAction.bas` carries that intent into `src/Core/Modules/modAuthUserCommands.bas`, which enforces current authorization and scope before persistence.
- The harness changes are under `tests/tooling`. `Test-Slice4beConfigCommands.ps1` dispatches the focused run. `Slice4beAdminCreateWarehouseOwner.ps1` and `Slice4beAdminCreateEdges.ps1` supply disposable private input to existing Create fixtures. `Slice4beAdminCreateRecording.ps1` checks emitted activity on rejected and successful entry. `Slice4beInitialAdminOnboarding.ps1` covers recovery and role scope through actual handlers. `Slice4beOnboardingGuides.ps1` exercises real recorded guide actions. `Slice4beOnboardingPeer.ps1` signs in independently, revokes, then cold-reopens.

## Decisions and constraints

The source administrator remains bound to the source. The destination Admin receives its two exact ordinary defaults. Existing credentials and guide permissions require no migration. Legacy six-choice saves omit the guide choice. Explicit guide changes require a current authorized administrator and the selected exact user, warehouse and station. Scoped recovery preserves unrelated arrivals and pre-existing contents. Uncertain cleanup reports failure with the incomplete destination ineligible for publication.

## Evidence and traceability

- The unchanged installed candidate passed 10/15 and 12/20 in the two baseline cohorts. `EmptyPin.RefusesBeforeBootstrap`, `Mismatch.NoDestination`, `ExistingHub.IncompleteDestinationNotPublishable`, `ExistingHub.OnlyOriginalEntriesRemain`, `ExistingHub.OrdinaryRetryCreatesWarehouse` and `Users.ExplicitGrant` exposed the missing contracts.
- `GuideTransfer.NextActionReadsRevocation` failed on v1 because the first Excel process reused cached authority. On v2 the exact-station display, `UncertainCleanup.UnrelatedArrivalPreserved` and `AccountProvisioning.NoPublishableFailureAfterCredentialPersistence` failed. On v3 each `Changed{User,Warehouse,Station}ExplainsPendingChoice` failed while the persistence guards passed.
- The final focused runs passed 59/59 and 39/39. Both private controls were masked, in bounds and consecutive in tab order. Empty/mismatched UI and typed Core requests and cancellation made no destination. The writer-failure cases preserved unrelated bytes and refused publication. A verified retry created a usable warehouse. Source identity and source Auth bytes survived. A normal destination Sign In confirmed exactly `ADMIN_MAINT` and `RECEIVE_POST`. Legacy saves preserved grant present and absent. Replaced, signed-out, read-only and rebound contexts and changed selection could not redirect an update.
- The new Admin authored through a recording, edited version 2, exported and imported. After a second signed-in Excel session saved revocation, the open editor refused guarded Remove/Save and transfer. Dirty Cancel and window X each discarded with zero Core WriteText attempts, unchanged guide bytes and no publication. A third process signed in normally and observed persisted revocation. Textual evidence and activity pairs were checked for private input. The evidence stores only outcomes.
- The v4 regressions passed Create 110/110, Auth and Users & Roles 259/259, Create recording 175/175, Release 1 full chain 33/33, guide saved-workbook restart 33/33, published/guide native regression 198/198, live roles 48/48. The two-row Published Guides search updated in 22–28 ms, and a no-match query in 76 ms. In the 31-row picker, open took 2.044 s, native X took 657 ms, and typing performed zero extra Core/library/policy reads. Scale was 35/36 because the separately scoped empty-search explanation is missing.
- The same pinned packages passed the reusable Production launcher checks 2/2 and ordered Receiving B0 replay 121/121. On the designated network share, 31 authorized rows opened in 6.325 s, keys and clear took 25–35 ms, no match took 76 ms, and reopened X disposed the window and form in 3.729 s. The NAS scale result was also 35/36 with only the separately scoped empty-search explanation open.
- Full NAS curation passed 100/101 with the same empty-search finding. Two successive single-click cycles opened and closed both Published Guides and Choose Tracked Actions. On this three-row flow the picker took 7.83–7.99 s to open. Creating from selected actions gave first feedback in 173 ms, displayed the editor in 3.636 s, and finished loading in 5.952 s.
- The saved-recording entry passed 40/40 on the same share. Its own Create Guide button needed one physical click, with feedback in 188 ms, the editor visible in 3.122 s, and loading complete in 6.594 s.
- Last verified November 24 after the final NAS run. All final cohorts closed Excel naturally and restored the invSys test settings. All 15 SHA-256 comparisons passed across the candidate, isolated verification copies and previously installed review build. Windows Application events 1000, 1001 and 1002 contained zero Excel faults over the final compile-to-last-test interval. The feature worktree retains unrelated untracked reports, the primary code worktree retains its local instructions, and the docs worktree retains earlier unrelated changes. No Excel process remained. Portable evidence is committed at code `ca190e4a` in `tests/integration/plan022_slice4be_operator_review.md` and `tests/integration/plan022_slice4be_remaining_acceptance.md`.
- Both static inventories have 392 components. Procedures increased from 6,779 to 6,796 and lines from 146,731 to 146,999. The longest new procedure is 29 lines. Duplicate-body candidates remain 193. Literal and unresolved Application.Run counts remain 9 and 45. The five added scanner candidates are native Users & Roles event roots.

| Candidate XLAM | SHA-256 |
| --- | --- |
| Admin | `779E00E0B43B978EF1805253FF7FD64C2844BEF49E346C713B28571526565032` |
| Core | `001B29415D3D4A4E07A32D1CAE943F9BEE72C93333B87964584D5724ACF217C7` |
| Designs.Domain | `637C888C02090A931D3776979ED092075BDA4C5004A6D930CD539790F28B6126` |
| Inventory.Domain | `2CE45AD045D058E88785D763C6AC9ED49BA1EC7D7CF0577D40FB8ED39021D984` |
| Operations | `41B80A9DE0B22CBE1E993D64C41AD7A285222BC3EC96903C77167CC7A7C295F3` |

## Do Not Repeat

An early peer adapter copied a stray conditional-compilation directive. Another instrumentation attempt duplicated the notice counter. Their compile diagnostics were attended and excluded from GREEN. The signed-out Users test raised “Tracking unavailable: the Users & Roles action could not be fully recorded.” Native dialogs were captured before acknowledgment. The final fixture observes this exact notice at the existing presentation boundary. Windows PowerShell's `Get-ChildItem -LiteralPath ... -Include` returned binaries in a textual privacy scan. Filtering actual extensions fixed that harness fault. The old direct Create regression omitted the new private arguments. Its success assertions passed after the fixture supplied them.

## Assumptions to re-verify

Recheck the candidate and installed hashes, Excel ownership, active desktop and designated Training location before a new operator session. Revalidate on any changed package set.

## Open questions and blockers

Justin's visible verdict on the ordinary onboarding path remains outstanding. The picker no-match status, catalog 49 Delete Item recording, Edit Save Item coverage and remaining adjacent work retain their separate status. No merge or deployment promotion was authorized.

## Immediate next action

Run the visible operator comparison on the pinned candidate in a disposable Training warehouse and record Justin's verdict.

## Critical references

`invSys-Design-v4.11.md` D8-CREATE-01, D8-USERS-01, D18-GUIDE-INPUT-01 and D18-SEARCH-INPUT-02; Plan 022; `invSys-Controls-v1.md`; `tests/integration/plan022_slice4be_operator_review.md`; local-only `reports/runtime/onboarding-v4-qualification` and `onboarding-v4-static`.
