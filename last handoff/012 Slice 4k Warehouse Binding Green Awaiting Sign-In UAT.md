# Slice 4k Warehouse Binding Green Awaiting Sign-In UAT

## 1. Goal and release outcome

Restore the selected-warehouse/session contract so the user can bind
`WHT7025AE` to computer station `X1-PRO-AI`, sign in to that warehouse, and
begin the visible Release 1 Receiving → Production → Boxing → Shipping test.
This advances v4.11 D12-D14 and Plan 022's dedicated-NAS acceptance checkpoint.

## 2. Current verified state

- Last verified: 2026-08-17.
- Code: `main`, commit `1d1387b` (`Fix warehouse target binding for computer stations`).
- Docs contract: `main`, commit `d5d3d3c` (`Document selected warehouse binding contract`).
- Active slice: Plan 022 Slice 4k, automated gates complete; visible sign-in and
  workflow acceptance remain.
- Excel is closed.
- User-owned unstaged changes remain in code `AGENTS.md` and docs
  `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`; they were preserved and not
  committed.

## 3. Decisions and constraints

- Ordinary Station is exactly the Windows computer name.
- A valid selected warehouse may enroll the exact current computer when its
  config still contains only a legacy station; arbitrary station names remain
  rejected.
- Target selection is atomic. A failed selection retains the prior target; a
  changed warehouse/station/root signs out the prior invSys session.
- The user must use the existing credential belonging to the selected
  warehouse. No human PIN/password was read, reset, or provisioned in this
  slice.

## 4. Evidence and traceability

- Symptom → root cause: both targets rejected and `<roaming>` remained →
  current computer absent from legacy `tblStationConfig` → Plan 022 Slice 4k →
  current-computer and same-action ribbon tests.
- Focused RED: 0/1 after a valid fixture reached legacy-station rejection.
- Focused GREEN: 2/2; Core target/auth/write regression: 28/28.
- Packaged XLAM: 74/74; packaged Ribbon: 142/142.
- Dedicated NAS: 12/12 in two clean sessions for `WHT7025AE` / `X1-PRO-AI`.
- Isolated ordered full chain: 30/30 through restart/reconciliation.
- Tool B redaction/non-mutation: 62/62; deterministic static baseline: 19/19.
- Static baseline: 150 components, 4,619 procedures, literal
  `Application.Run` 9, unresolved expressions 48, duplicate bodies 184.

## 5. Do Not Repeat

- Do not restore hard-coded station `S1` in Plan 022 NAS validation.
- Do not report `GetConnectionStatus()` as the failed-attempt reason; it
  intentionally describes the prior committed target after failure.
- Do not carry a signed-in session across a changed target identity.

## 6. Assumptions to Re-verify

- The existing `Justin` user and credential are present in the WHT Auth
  workbook; this was not tested because the credential is user-held.
- Re-verify the user's capability rows at `X1-PRO-AI` when role actions are
  exercised. Selection and ordinary sign-in do not require a role capability.
- Visible form sizing and complete dedicated-NAS business workflow remain user
  acceptance evidence despite automated geometry and isolated-chain success.

## 7. Open questions and blockers

- No automated blocker remains.
- User acceptance must confirm the ribbon changes to `WHT7025AE`, Sign In shows
  `WHT7025AE` / `X1-PRO-AI`, and the existing WHT credential is accepted.
- If credential authentication succeeds but role buttons remain unavailable,
  inspect/provision the existing user's current-station capability rows without
  changing the PIN.

## 8. Immediate next action

Open Excel, connect the server, select `WHT7025AE`, verify the ribbon and Sign
In target both show `WHT7025AE` / `X1-PRO-AI`, then sign in with the existing
WHT credential.

## 9. Critical references

- `src/Core/Modules/modNasConnection.bas`
- `src/Core/Modules/modRibbonRuntimeStatus.bas`
- `src/Core/Modules/modAuth.bas`
- `tests/integration/plan022_slice4k_target_binding_red_results.md`
- `tests/integration/plan022_slice4k_target_binding_green_results.md`
- `tools/validate_plan022_nas_runtime.ps1`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4k
