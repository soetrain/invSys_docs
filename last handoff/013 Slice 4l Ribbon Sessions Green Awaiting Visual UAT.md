# Slice 4l Ribbon Sessions Green Awaiting Visual UAT

## 1. Goal and release outcome

Make Operations/Admin session controls accurately reflect warehouse selection
and separate the Windows/NAS server layer from invSys authentication so the
visible Release 1 workflow can begin without ambiguous or stale Ribbon state.
This advances v4.11 D-NAS and Plan 022's dedicated-NAS acceptance checkpoint.

## 2. Current verified state

- Last verified: 2026-08-17.
- Code: `main`, commit `c0e4df8` (`Fix ribbon session state and server sign out`).
- Docs contract: `main`, commit `04ea8ff` (`Define explicit server and invSys session controls`).
- Active slice: Plan 022 Slice 4l, automated gates complete; visible Ribbon
  acceptance remains.
- Excel is closed.
- User-owned unstaged changes remain in code `AGENTS.md` and the NAS/server
  pointer at the top of docs `invSys-Design-v4.11.md`; neither was committed.

## 3. Decisions and constraints

- **invSys Sign In / invSys Sign Out** controls the invSys identity and
  capability session. invSys Sign Out retains server access for user switching.
- **Server Sign In / Server Sign Out** controls the Windows/NAS session. Server
  Sign Out also signs out invSys, clears the target, disconnects the SMB share,
  and disables capability-gated operator controls.
- Disconnected invSys Sign In must not revive remembered target state or open
  the credential form; it gives connect-first guidance.
- Send To must invalidate the deployed Operations IDs immediately; Runtime
  Context is not a refresh mechanism.

## 4. Evidence and traceability

- Stale server label → retired IDs invalidated instead of live Operations IDs →
  v4.11 D-NAS / Plan 022 Slice 4l → exact action/source and dedicated-NAS tests.
- Initial focused RED: 0/8; added disconnected access-status RED: 8/9.
- Focused GREEN: 9/9; same-action Excel: 2/2; Core session regression: 30/30.
- Ribbon generation: 48/48; packaged XLAM: 74/74; packaged Ribbon: 140/140.
- Dedicated NAS: 16/16 across two clean sessions, including immediate selected
  status, real SMB disconnect, signed-out labels/access, and reconnect.
- Isolated ordered full chain: 30/30 through restart/reconciliation.
- Plan 022 launchers: 24/24; Slice 4j: 18/18; static baseline: 19/19.
- Static metrics: 150 components, 4,630 procedures, literal `Application.Run`
  8, unresolved expressions 47, duplicate bodies 184. Dynamic calls improved
  from 9/48 to 8/47.

## 5. Do Not Repeat

- Do not reintroduce generic Sign In/Sign Out/Connect Server labels.
- Do not invalidate only retired role-specific Ribbon IDs after Operations
  Send To selection.
- Do not let invSys Sign In call remembered-target resolution before confirming
  a live server session.
- Do not implement same-project `modAuth.SignOut` through `Application.Run`;
  `ClearWarehouseTarget` now uses a direct typed call.

## 6. Assumptions to Re-verify

- Visible Excel Ribbon rendering must show the new labels after a clean Excel
  restart and updated add-in load.
- The user's existing WHT credential remains user-held and was not changed.
- Role capability scope for the existing user is still checked when each role
  control is exercised.

## 7. Open questions and blockers

- No automated blocker remains.
- User acceptance must confirm immediate WHT status after Send To; Server and
  invSys toggle labels; disabled role controls and connect-first prompt after
  Server Sign Out; then the full business workflow.

## 8. Immediate next action

Open Excel cleanly, use Server Sign In, select `WHT7025AE`, verify immediate
status, authenticate with invSys Sign In, then exercise both Sign Out paths
before continuing the full workflow.

## 9. Critical references

- `src/Core/Modules/modRibbonRuntimeStatus.bas`
- `src/Core/Modules/modRoleEventWriter.bas`
- `src/Core/Modules/modNasConnection.bas`
- `tools/build-xlam.ps1`
- `tests/tooling/Test-Plan022Slice4lRibbonSessionControls.ps1`
- `tests/integration/plan022_slice4l_ribbon_session_red_results.md`
- `tests/integration/plan022_slice4l_ribbon_session_green_results.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`, D-NAS operator workflow
- `0 plan docs/xlam_invSys/D-NAS_Procedure_Contract.md`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4l
