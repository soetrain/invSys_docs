# Slice 4m Admin Station Capability Green Awaiting Visual UAT

## 1. Goal and release outcome

Allow the existing `WHT7025AE` administrator to authenticate at the selected
computer station without weakening exact capability enforcement, so the
visible Release 1 workflow can resume. This advances v4.11 D-NAS and Plan 022
Slice 4 dedicated-NAS acceptance.

## 2. Current verified state

- Last verified: 2026-08-17.
- Code: `main`, commit `b99021a` (`Fix Admin sign-in station capability transition`).
- Docs contract: `main`, commit `94b21e8` (`Define legacy station capability transition`).
- Active slice: Plan 022 Slice 4m; automated gates complete, visible Admin
  sign-in acceptance remains.
- Excel is closed and the rebuilt five-package set is in `deploy/current`.
- User-owned unstaged changes remain in code `AGENTS.md` and the NAS/server
  pointer at the top of docs `invSys-Design-v4.11.md`; neither was committed.
- During the user-authorized Excel close, the unsaved
  `WHT7025AE.Receiving.Operator.xlsm` session was closed without saving. No
  canonical warehouse workbook was changed by that close operation.

## 3. Decisions and constraints

- `StationId` remains exact station scope; `S1` is not a wildcard or permanent
  runtime alias.
- Only after an exact user's invSys secret validates may Core copy that same
  user's effective active legacy `S1` capabilities to the selected station
  when it exactly matches the Windows computer name.
- The transition preserves warehouse scope and validity dates, retains source
  rows, does not overwrite an existing current-station row, respects target
  denies, and cannot invent a capability absent from effective `S1` scope.
- A required capability missing at `S1` still returns
  `AUTH_NO_CAPABILITIES` and causes no transition.

## 4. Evidence and traceability

- Symptom: Admin sign-in showed `WHT7025AE` / `X1-PRO-AI` then **User lacks
  the required capability**.
- Root cause: credential resolution was correct, but `ADMIN_MAINT` remained at
  generated placeholder station `S1` while the target now used the computer
  name.
- Focused RED: 1/2; legacy Admin transition failed while missing-role security
  passed.
- Focused GREEN: 3/3, including explicit target-deny preservation.
- Core NAS/target/auth/session: 33/33.
- Packaged XLAM: 74/74; packaged Ribbon: 140/140.
- Dedicated `WHT7025AE` NAS: 16/16 using the automation user; human UAT PIN
  hash was not read or changed.
- Ordered full chain: clean-session 30/30. An initial 28/29 run passed every
  business/restart check but its final extractor hit an orphaned Excel process;
  the orphan was terminated before the complete rerun.
- Plan 022 launcher contracts: 24/24; packaged no-eligible state: 3/3; Slice
  4j: 18/18; Slice 4l: 9/9.
- Static: expected baseline drift 13/19, regenerated deterministic GREEN
  19/19. Metrics: 150 components, 4,633 procedures, 8 literal
  `Application.Run` targets, 47 unresolved expressions, 184 duplicate bodies.
- Detailed evidence:
  `tests/integration/plan022_slice4m_admin_station_capability_results.md`.

## 5. Do Not Repeat

- Do not bypass station scope or treat `S1` as a match for every computer.
- Do not transition capabilities during Server Sign In or target selection;
  those actions prove storage access, not invSys identity.
- Do not overwrite explicit current-station rows, including `DENY`.
- Do not run a package build while Excel has invSys add-ins or operator
  workbooks open.

## 6. Assumptions to Re-verify

- The existing human user still has an effective `ADMIN_MAINT` row at `S1` in
  the selected WHT Auth workbook.
- The account-scoped add-in registration still points to `deploy/current`.
- The visible Admin sign-in will create the exact `X1-PRO-AI` capability rows
  on first successful credential submission.

## 7. Open questions and blockers

- No automated blocker remains.
- Human credential input is intentionally unavailable to automation; visible
  Admin sign-in is required to close this checkpoint.

## 8. Immediate next action

Open Excel cleanly, select `WHT7025AE`, use Admin **invSys Sign In** with the
existing warehouse user, and confirm the Admin controls enable at
`X1-PRO-AI` without the capability error.

## 9. Critical references

- `src/Core/Modules/modAuth.bas`
- `src/Core/Forms/frmSignIn.frm`, `mBtnOK_Click`
- `tests/unit/TestPhase6CoreSurfaces.bas`
- `tests/integration/plan022_slice4m_admin_station_capability_results.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`, D-NAS item 8
- `0 plan docs/xlam_invSys/D-NAS_Procedure_Contract.md`
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, version 1.10
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4m
