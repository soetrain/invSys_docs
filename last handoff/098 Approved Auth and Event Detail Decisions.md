# Slice 4be: approved Auth and Event Detail decisions

## 1. Goal and release outcome

Complete Release 1/Slice 4be under Architecture v4.11 and Plan 022 while retaining
accepted workflows. The user approved both pending proposals and asked whether
Goal execution is ready to resume. This checkpoint records approval, not runtime
implementation or release completion. The Goal tool still reports blocked;
resumption is controlled by the user, not by an agent status update.

## 2. Current verified state

Last verified: 2026-09-24. Both repositories are on main. Code approval checklist
commit **5ade584** is pushed; docs base is **c0f5cac**, with the approval changes
and this handoff included in the ensuing documentation commit. Runtime remains
**9cf03db**, candidate `deploy/validation-production-diagnostics`, with prior
390/390 focused GREEN retained. No runtime/package changes or behavioral reruns
occurred in this turn. Controls version is **1.240**.

Preserved unrelated changes: frmEventDetail.frm's 18 measurement lines, handoff067,
and untracked critique023. No Excel process was present at desktop preflight.

## 3. Decisions and constraints

The user's explicit 2026-09-24 approval makes both normative decisions effective:
D8-A ordinary Auth reads validate existing exact-target authority without
creation/repair/saves, while explicit authorized provisioning retains setup;
D18 Event Detail may use Locked=False solely for non-editable selection/scrolling.
Architecture, Plan 022, controls and the remaining-acceptance checklist are
synchronized. Historical pending wording is superseded, not historical results.
Multiline readability remains separate and open. No approval was inferred for
carrier authority or unresolved guide-transfer details. D12/D13 remain binding.

## 4. Evidence and traceability

At **16:31:13 UTC**, the read-only Win32 probe reports GetCursorPos success/error0,
readable UOI_IO=True, and zero Excel processes. No cursor movement, elevation,
permission changes, workbook operations or credential inspection occurred.
This is restored access at one instant, not proof of sustained visible-test access.
User reports changing host settings after host/client locking observations;
the precise cause of earlier error5 is still unproven.

Event Detail existing native-input RED: **42 PASS / one expected movement FAIL**,
with all 41 prior checks retained. See the detail-overflow result record.
D8-A still requires focused packaged Core caller coverage beyond the seven
existing Shipping Auth recreation failures. Documentation diffs and whitespace
checks pass; approval recording itself supplies no new product RED/GREEN.

## 5. Do Not Repeat

Do not ask again for either recorded approval. Do not treat approval as acceptance,
earlier historical pending text as a new blocker, or the successful short desktop
probe as sustained-input proof. Preserve the current runtime candidate and user
edits. Handoffs096/097 retain earlier evidence and coverage accounting only.

## 6. Assumptions to Re-verify

Recheck desktop access, Excel closure and frozen candidate hashes before a gate.
Do not rebuild while relevant workbooks/add-ins are open. Run one Excel gate at
a time using disposable fixtures and retain the original failure evidence.

## 7. Open questions and blockers

The two architecture approval blockers are removed. Implementation, native visible
evidence, multiline readability, comprehensive control coverage, comparison,
guide transfer, carrier authority and broader Release 1 gates remain open.
The Production candidate still needs its scoped visible/owner/full-chain gates.

## 8. Immediate next action

Revalidate the existing native-scroll RED against the frozen candidate, then
implement the approved Event Detail lock change and prove packaged GREEN and
unchanged values/source bytes before broadening regression coverage.

## 9. Critical references

- Architecture v4.11: approved D8-A and Event Detail scrolling decision headings.
- Plan 022 current decision update; invSys-Controls-v1.md version1.240.
- Code tests/integration/plan022_slice4be_remaining_acceptance.md.
- Code tests/integration/plan022_slice4be_detail_overflow_results.md.
- Code tests/tooling/Slice4beViewerEventDetail.ps1.
- Handoffs096/097 for preserved Production evidence, wrappers and source census.
