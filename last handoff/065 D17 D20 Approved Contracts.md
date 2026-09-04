# D17–D20 Approved Contracts

## Goal and release outcome

Advance Release 1 from accepted deployment to approved multi-server advisory
aggregation, curated Viewer Action Paths, controlled Admin Viewer projection,
archive-first retention, and owned save-feedback behavior.

## Current verified state

- Code `main` remains `940e1c7`; no implementation source changed in this
  specification turn.
- Docs `main` is pushed at `002701a` (`Approve multi-server aggregation and
  curated Action Paths`).
- Architecture v4.11 now contains D17 multi-server source set, D18 curated
  Action Paths, D19 Admin Viewer-event/retention, and D20 notification
  ownership. Plan 022 and controls v1.52 are synchronized.
- Slice 4bf deployment is physically accepted. Production remains user accepted.
- Existing single-source Admin aggregator was visibly used successfully: it
  read one published snapshot and wrote an advisory result. That is correct for
  the old contract; two-server selection is not yet implemented.
- Code worktree has unrelated generated reports and operational inbox files;
  preserve them. Docs worktree was clean after `002701a`.

## Decisions and constraints

- GitHub XLAM download/repository cloning is developer-only. Operator stations
  consume a verified D16 SharePoint/NAS feed/cache only.
- D17 source set is `ADMIN_MAINT`, explicit and session-scoped. It reads only
  selected published snapshots, never changes Send To, creates runtimes, or
  stores NAS credentials.
- D18 Action Path is a versioned curated human training record, not macro
  automation, control telemetry, or audit history. `ACTION_PATH_MAINT` creates/
  imports/exports; signed-in Viewer users search/read published paths.
- D19 Admin Viewer-event setting never suppresses canonical audit/security or
  authority events. Retention is archive-first and disabled by default.
- D20 forbids invSys changing Windows/NAS-client/global Office notifications;
  diagnose notification owner before changing invSys feedback.
- Slice 4bg remains not approved.

## Evidence and traceability

- `src/Admin/Modules/modAdmin.bas:Admin_AggregateGlobalSnapshot_Click`
- `src/Admin/Modules/modAdminConsole.bas:RunHQAggregationFromAdmin`
- `src/Core/Modules/modHqAggregator.bas:RunHQAggregation`
- `tests/tooling/Test-Slice4bdAdminAggregatorCommand.ps1`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`

## Do Not Repeat

- Do not force multi-server behavior into the current single-target aggregator.
- Do not store NAS credentials in selected aggregation source records.
- Do not treat the pasted Windows-toast hypotheses as root cause without a
  controlled invSys/OS owner diagnosis.

## Immediate next action

Create and run a focused packaged Admin source-set RED test, then implement
D17 discovery/selection without changing current warehouse targeting.
