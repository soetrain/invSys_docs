# Slice 4bd NAS Deployed — Physical UAT Pending

## Goal and release outcome

Advance Release 1 Architecture v4.11 / Plan 022 Slice 4bd with a session-scoped multi-server advisory Aggregator source set, preserving warehouse-local authority and preparing physical proof on two NAS-backed warehouses and computers.

## Current verified state

- Code `main`: `e06e4ef`, pushed. Docs `main`: `4c30869`, pushed.
- NAS five-package release `R1-20260903-e06e4ef` was published and the station updater reported `APPLIED` on 2026-09-03 with Excel closed.
- Slice 4bd isolated gates GREEN: source-set contract 6/6; Admin command contract 3/3; isolated Core/Admin build; explicit two-source WarehouseId/System_Key aggregation; packaged Admin form initialization.
- Static-maintenance evidence was regenerated to a temporary local folder. Do not commit generated evidence.
- User/runtime changes remain dirty in the code checkout, including generated `deploy/current` XLAMs, runtime reports, inbox workbooks, and `DesignsDomain/`. Do not stage or overwrite them. Excel must be closed before any next build/deployment.

## Decisions and constraints

- D17 is locked. **Aggregate Global Snapshot** opens `frmAggregationSources` under `ADMIN_MAINT`; source selection is session-only and must not change Send To, authentication, runtime creation, or inventory authority.
- Discovery requires readable `tblWarehouseConfig`, its configured published snapshot, and matching snapshot WarehouseId. Metadata is read-only with macros disabled; credentials are cleared after Windows connection and never retained in source/config/audit/event/log.
- Admin serializes selected file paths as primitives for Core; no VBA object crosses the XLAM boundary. Core reads only listed files and remains headless.
- GitHub is developer-only. NAS is the approved operational update feed; SharePoint updater work is deferred.

## Evidence and traceability

- `tests/tooling/Test-Slice4bdMultiServerSourceSet.ps1` — 6/6 GREEN.
- `tests/tooling/Test-Slice4bdAdminAggregatorCommand.ps1` — 3/3 GREEN.
- `tests/tooling/Test-Slice4bdExplicitSourceAggregation.ps1` — selected source identity plus WarehouseId/System_Key GREEN.
- `tests/tooling/Test-Slice4bdAdminFormSmoke.ps1` — packaged form initialization GREEN.
- `tools/build-xlam.ps1` — isolated Core/Admin and current five-package builds GREEN.

## Do Not Repeat

- Do not restore the former immediate single-root scan.
- Do not pass a VBA `Collection` between Admin and Core XLAMs.
- Do not call the isolated proof physical two-warehouse acceptance.

## Assumptions to Re-verify

- NAS connections and two distinct warehouse runtime/config roots are accessible when UAT begins.
- Fresh Excel loads the registered release.

## Open questions and blockers

- Physical two-computer/two-warehouse UAT remains and needs an operator with two NAS roots and two Excel stations.

## Immediate next action

Open fresh Excel on the first station, use Admin > Aggregate Global Snapshot, discover/select both configured warehouse snapshots, and record source state and advisory output before complementary-station proof.

## Critical references

- `src/Admin/Forms/frmAggregationSources.frm`
- `src/Admin/Modules/modAdmin.bas` — `Admin_AggregateGlobalSnapshot_Click`
- `src/Admin/Modules/modAdminConsole.bas` — source discovery/action
- `src/Core/Modules/modHqAggregator.bas` — `GenerateGlobalSnapshotFromFiles`
- Plan 022 and Architecture D17
