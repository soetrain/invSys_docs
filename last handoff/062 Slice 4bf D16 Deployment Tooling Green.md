# Slice 4bf D16 Deployment Tooling Green

## Goal and release outcome

Implement Architecture v4.11 D16's immutable five-package publication,
automatic station update, and local-admin rollback without changing inventory
or Domain authority. Slice 4bf tooling is GREEN; physical SharePoint/NAS/station
UAT remains required for Release 1 acceptance.

## Current verified state

- Code `main` is pushed at `ad53c52` (`Implement D16 immutable release deployment tooling`).
- Docs `main` is pushed at `ad7e19d` (`Lock D16 deployment and rollback contract`).
- D16 is normative in `0 plan docs/xlam_invSys/invSys-Design-v4.11.md`; Plan
  022 and controls are synchronized.
- Focused D13 test: RED 0/9 before implementation; GREEN 16/16 on 2026-09-03.
  Slice 13 five-package cutover regression is GREEN 14/14.
- Excel was closed for testing; re-check before any package build, registration,
  publish, task installation, or rollback.
- Code worktree has unrelated generated reports and operational inbox workbooks
  still uncommitted; do not stage or delete them. Docs worktree was clean after
  the D16 commit.

## Decisions and constraints

- D16 release feed: `<PathSharePointRoot>\Addins\Releases\<ReleaseId>` plus
  `current-release.json`; immutable `R1-5` five-package manifest with SHA-256.
- GitHub is source/review authority. NAS may mirror the feed only outside all
  warehouse authority/data/config/auth/inbox/outbox/operator paths.
- Task is logon + 15 minute cadence, automatic/no prompt, and defers whenever
  Excel is open. Cache retains three releases. Only Operations/Admin are Excel
  startup leaves; Core/Domain stay headless.
- Rollback requires closed Excel, local Windows administrator, a retained
  verified release, `-ConfirmRollback`, and a constrained non-sensitive reason
  code. It is package-only.
- Slice 4bg is explicitly not approved; do not implement it.

## Evidence and traceability

- `tests/tooling/Test-Slice4bfDeployment.ps1`: immutable manifest, Excel
  deferral, hash verification, leaf-only registration, third-party add-in and
  authority-workbook non-mutation, non-admin rollback gate, tamper rejection,
  and injected registration-failure restoration.
- `tests/tooling/Test-Slice13OperationsCutover.ps1`: 14/14 GREEN.
- `tools/validate_phase6_packaged_xlams.ps1` was invoked against unchanged
  `deploy/current`; it returned without an error or reported failure. The static
  maintenance scan was stopped because it did not complete its output artifact
  in this environment; no VBA source changed in Slice 4bf, and previously
  committed static evidence remains the applicable baseline.

## Do Not Repeat

- Do not use the legacy in-place `deploy_current_xlams_to_nas.ps1` path for D16
  release publication.
- Do not remove all Excel `OPEN*` entries: the updated registration script
  removes only invSys entries and preserves third-party add-ins.
- Do not publish or install the scheduled task until the real configured
  SharePoint/NAS Addins root is identified and the user authorizes that external
  deployment action.

## Assumptions to Re-verify

- A real writable `<PathSharePointRoot>\Addins` target and optional NAS mirror
  have not been selected or populated by this slice.
- The station on which the scheduled task/rollback will run has a local Windows
  administrator available.
- The pending two-PC/two-NAS-warehouse physical Aggregator UAT remains open.

## Open questions and blockers

- Slice 4bf physical UAT needs a chosen real feed root, a release ID, and a
  station where the task can be installed. Those are external-state choices;
  code must not guess or create warehouse paths.
- Slice 4be is approved, but the durable captured-vs-derived Event Viewer Action
  Path design decision was not explicitly answered before D16 work began.

## Immediate next action

With the user, choose the real D16 Addins feed root and execute the focused
physical publish/update/task proof with Excel closed and a local administrator.

## Critical references

- `tools/publish_invsys_release.ps1`
- `tools/update_invsys_station.ps1`
- `tools/register_invsys_update_task.ps1`
- `tools/rollback_invsys_station_release.ps1`
- `tools/register_current_addins.ps1`
- `tests/tooling/Test-Slice4bfDeployment.ps1`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`
