# Slice 4bf Physical Deployment Preflight

## Goal and release outcome

Complete D16 immutable five-package deployment without allowing Git, a
warehouse runtime root, or inventory authority to become a station dependency.
The physical NAS-feed publication and this-station registration are complete;
scheduled-task elevation and visible next-Excel verification remain.

## Current verified state

- Code `main` is pushed at `940e1c7` (`Remove Git checkout dependency from
  station updater task`); prior D16 implementation is `ad53c52`.
- Docs `main` is pushed at `fca3add` (`Record D16 physical deployment preflight`).
- `R1-20260903-ad53c52` is published as a hash-verified immutable R1-5 release
  to the separate NAS deployment share's `Addins` feed. No canonical warehouse
  runtime root was used.
- With Excel closed, this station cached that release and changed only the
  Operations/Admin startup entries to its local cached leaf XLAMs. The local
  station-maintenance agent is present and hash-verified.
- No scheduled task exists: this non-elevated session received Windows access
  denial when creating it. No Excel process was left running.
- D13 focused deployment test is GREEN 18/18; Slice 13 cutover remains GREEN
  14/14. The direct `powershell -File` updater invocation and local-agent task
  action are now explicitly protected.
- Code worktree retains unrelated generated reports and operational inbox files;
  do not stage/delete them. Docs worktree was clean after the physical-preflight
  record commit.

## Decisions and constraints

- The live warehouse `PathSharePointRoot` is a stale/unreachable synced path on
  this account; do not edit its configuration or invent a replacement. The
  separate NAS deployment share is the valid D16 mirror used for this proof.
- The scheduled task must invoke its local hash-verified station agent under
  the user's local invSys deployment directory, never the repository checkout.
- Task Scheduler registration is subject to Windows local policy. This station
  requires an elevated Windows session; credentials must never be read from
  files, source, tests, logs, or handoffs.
- The task installer normally stages the agent before it asks Task Scheduler to
  register. It retains logon + 15-minute cadence and Excel-open deferral.
- Slice 4bg remains explicitly not approved.

## Evidence and traceability

- `tests/tooling/Test-Slice4bfDeployment.ps1`: GREEN 18/18.
- `tools/publish_invsys_release.ps1`: physical immutable release publication.
- `tools/update_invsys_station.ps1`: physical local-cache/leaf-registration
  proof.
- `tools/register_invsys_update_task.ps1`: agent staged; task registration
  blocked only by Windows access control; read-only check confirmed no task.

## Do Not Repeat

- Do not publish to the canonical warehouse runtime or its legacy flat XLAM
  files. The separate D16 `Addins\Releases` feed is authoritative for this
  deployment proof.
- Do not use a Git checkout as the scheduled task's updater path.
- Do not attempt to bypass Windows elevation or inspect credentials.

## Assumptions to Re-verify

- The NAS deployment feed remains reachable from the elevated user session.
- Excel is closed immediately before installing the task and before visible
  startup verification.

## Open questions and blockers

- User-side Windows elevation is required to register `invSys.StationUpdate`.
- Physical acceptance still needs opening Excel after successful task
  installation and confirming the cached Operations/Admin packages load
  normally.
- The two-PC/two-NAS-warehouse Aggregator UAT and approved Slice 4be Viewer
  work remain separate open release items.

## Immediate next action

From an elevated PowerShell session with Excel closed, run the existing D16
task installer against the established NAS Addins feed, then open Excel and
confirm the cached Operations/Admin add-ins load visibly.

## Critical references

- `tools/register_invsys_update_task.ps1`
- `tools/update_invsys_station.ps1`
- `tools/publish_invsys_release.ps1`
- `tests/tooling/Test-Slice4bfDeployment.ps1`
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md` (D16)
