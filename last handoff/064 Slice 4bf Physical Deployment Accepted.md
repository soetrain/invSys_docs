# Slice 4bf Physical Deployment Accepted

## Goal and release outcome

Prove D16 immutable five-package deployment, automatic station update, and
rollback-ready registration without depending on Git or a warehouse runtime
root. Slice 4bf physical deployment acceptance passed on 2026-09-03.

## Current verified state

- Code `main` is pushed at `940e1c7`; D16 tooling implementation is `ad53c52`.
- Docs `main` is pushed at `403225b` (`Record D16 physical deployment acceptance`).
- The separate NAS deployment feed contains immutable release
  `R1-20260903-ad53c52`; no canonical warehouse runtime root was used.
- The station hash-verified/cached that release and registered only cached
  Operations/Admin leaf XLAMs. Git is not a workstation dependency.
- `invSys.StationUpdate` is enabled with logon and 15-minute triggers and an
  existing local verified station agent. Its on-demand run returned `0x00000000`
  and `APPLIED` for the cached release.
- A clean Excel startup loaded cached Operations and Admin. The verification
  Excel instance was then closed.
- D13 deployment suite: GREEN 18/18. Slice 13 cutover regression: GREEN 14/14.
- Code worktree retains unrelated generated reports and operational inbox files;
  do not stage/delete them. Docs worktree was clean after acceptance commit.

## Decisions and constraints

- D16 feed uses a separate NAS deployment share as a mirror, never a warehouse
  runtime root. The configured synced SharePoint path remains unreachable on
  this account and must not be silently rewritten.
- Scheduled updater executes local hash-verified PowerShell station tooling,
  never a Git checkout. The five XLAM release remains exactly five packages.
- Rollback continues to require local Windows administrator, closed Excel, a
  retained verified release, and a non-sensitive reason code.
- Slice 4bg is not approved; do not implement it.

## Evidence and traceability

- `tests/tooling/Test-Slice4bfDeployment.ps1` (18/18 GREEN)
- `tools/publish_invsys_release.ps1`
- `tools/update_invsys_station.ps1`
- `tools/register_invsys_update_task.ps1`
- Plan 022 Slice 4bf physical deployment evidence, controls v1.51, D16.

## Do Not Repeat

- Do not deploy legacy flat XLAMs in the NAS root or use the canonical warehouse
  root for Addins.
- Do not make a workstation Git checkout a scheduler dependency.

## Assumptions to Re-verify

- The separate NAS deployment feed remains reachable and the elevated task
  continues to run under the intended Windows account.
- Excel is closed before future manual update/rollback work.

## Open questions and blockers

- Two-PC/two-NAS-warehouse Aggregator physical UAT remains open.
- Approved Slice 4be Viewer remains open; its Action Path durability design was
  not explicitly selected before deployment work began.

## Immediate next action

Resume the pending two-PC/two-warehouse Aggregator physical UAT, or establish
the approved Viewer Action Path contract before Slice 4be implementation.

## Critical references

- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md` (D16)
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`
- `tests/tooling/Test-Slice4bfDeployment.ps1`
