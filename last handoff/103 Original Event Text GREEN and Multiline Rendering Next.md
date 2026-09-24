# Slice 4be: original text GREEN; multiline rendering next

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11/Plan022 without regressing
accepted workflows. Goal remains active/incomplete. This checkpoint restores
D18 original event values through Operations/Admin Viewer and Event Detail.
Complete multiline visibility and the broader acceptance checklist remain open.

## 2. Current verified state

Last verified: 2026-09-24, after 20:33 UTC. Both repositories are on main.
Code **ad16164** is committed/pushed. Docs parent **899bea2**; the ensuing commit
contains Plan022, controls1.245 and this handoff. Verify its push/upstream state.
Preserve unrelated docs handoff067 edits and untracked critique023; both hashes
match their earlier private pins. Runtime reports/images remain ignored.

New isolated candidate: `deploy/validation-detail-original-text`. Accepted
deployment and predecessor `validation-auth-read-separated` are preserved.
Excel and both chain controllers are closed; settings and tracked generated
reports are restored. No running test should be resumed or restarted.

## 3. Decisions and constraints

D8-A and D18 scrolling remain approved. The decoder correction restores the
existing contract; no normative amendment is needed for it. Core wire escaping,
captured context, exact System_Key, unknown columns, profile filtering and
read-only/headless authority remain unchanged. Only the existing Viewer decoder
changes in runtime source; no new component, procedure or dynamic call.

Do not confuse retained multiline values with complete visual rendering. A
scrollable read-only Frame containing caption/value Labels for permitted
multiline fields of the selected contributing line is a tentative implementation
direction, not an approved new contract. It could preserve the current selection
model and avoid a text editor. Reconcile any refinement in the normative spec,
Plan and catalog before implementation; do not infer approval from this handoff.
Private planning note: `reports/runtime/detail-multiline-next-proposal.md`.

## 4. Evidence and traceability

Primary sanitized record:
`tests/integration/plan022_slice4be_original_text_results.md`.

- Packaged actual publication/Viewer/Detail RED **52 PASS / four expected FAIL**;
  GREEN **56/56**, exact identities and all preceding 50 retained. LiteralEscapes
  and MixedText fail in each role before correction; actual linebreak/tab cases
  already pass. Both runs close normally, restore settings and preserve hashes.
  Five instrumented compiles each; four principal GREEN captures reviewed.
- All five packages build/compile; Operations cold-start resolves locally. Of
  244 compiled components, only Operations/frmInventoryViewer changes.
- Viewer/filter/Shipping-state **98/98**, exact prior identities, three reviewed
  images, normal closure, settings/hash preservation and zero Application events.
- First chain: **5 PASS / one harness FAIL**; live-role **32 PASS / one harness
  FAIL**; Create Warehouse **15/15**. Excel crashes during canonical projection
  rebuild at modProcessor.RunBatchReportForAutomation, HRESULT0x800706BE.
  Events1000/1001 identify ntdll.dll/0xc0000028. Verified recovery Excel termination
  lets the original controller restore settings/reports. Cause remains unproven.
- Unchanged-candidate retry: **32/48/15**, exact prior identities, normal
  unassisted closure, settings/three reports restored, five packages/275 tooling
  pins preserved, zero Application failures. UTC20:26:53--20:32:19. The failed
  attempt remains recorded and is not converted into successful evidence.
- Static: 251 components, 6047 procedures, 132986 lines (+14 for the decoder),
  nine literal/45 unresolved calls, 191 duplicate groups, all28 oversized limits
  preserved. Three schemas pass; layout8/8 and7/7; 276 PowerShell files parse.

Private receipts use `reports/runtime/detail-original-text-`: `verification.json`,
`viewer-verification.json`, `compiled-delta.json`, `package-pins.json`,
`static-verification.json`, `chain-failure-verification.json` and
`chain-retry-verification.json`. Exact report roots are in the primary record.

## 5. Do Not Repeat

Do not replace escaped newline sequences before escaped backslashes. Do not
accept geometry/string equality alone as complete multiline visual evidence.
Do not edit pinned tests during a running gate or hash XLAMs while Excel holds
exclusive handles; a sharing violation is not proof of content mutation.
Do not reattach COM or issue a second Quit during original cleanup. Retain its
settings snapshot; identify any separate post-crash recovery process precisely.
Normalize both sides of process-time guards to UTC before comparison.

Do not blindly rerun the broader Shipping/Boxing route: its recorded on-screen
caption/z-order failure remains unresolved (1632 PASS / one capture exception,
82 unreached). See the preceding capture lifecycle evidence; no alternate raise
flags or runtime fix for that failure are implemented here.

## 6. Assumptions to re-verify

Confirm Excel closure and candidate pins before the next packaged gate. Desktop
capture succeeds in the focused gates; neither prior error5 nor the intermittent
native Excel crash has an established root cause. Earlier candidate evidence
retains its stated scope and does not establish comprehensive final acceptance.

## 7. Open questions and blockers

Complete multiline rendering, comprehensive Operations/Admin tracking (61 pending
Production buttons), guide transfer/provenance, comparison/presentation, carrier
D5 authority, seven owner capture corrections, broader capture/normal closure
and human/NAS acceptance remain. Consult the maintained remaining-acceptance
checklist; do not redefine the Release 1 goal around this decoder checkpoint.

## 8. Immediate next action

Reconcile the multiline rendering refinement with D18, then establish packaged
Viewer/Detail RED for complete multiline reachability while retaining all56 checks.

## 9. Critical references

- Architecture v4.11 D18/D8-A; Plan022; controls1.245.
- `tests/integration/plan022_slice4be_remaining_acceptance.md`.
- `tests/integration/plan022_slice4be_original_text_results.md`.
- `tests/integration/plan022_slice4be_capture_lifecycle_results.md`.
- `tests/tooling/Slice4beViewerEventDetail.ps1`, `Slice4bePublishedProjectionFixture.ps1`.
- `src/Operations/Forms/frmInventoryViewer.frm`: ViewerUnescape.
- `src/Operations/Forms/frmEventDetail.frm`: RefreshFields, FitFieldColumns, ClearContent.
- `src/Operations/ClassModules/cEventDetailController.cls`: Fields, ContextValid, Invalidate.
