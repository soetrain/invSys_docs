# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4ae restricts Production item search to assignment cells and makes Process row
IDs unique and stable across save/reopen.

# Current verified state

- Last verified 2026-08-25: code `main` at `8128457`; normative/plan/control
  docs at `cedb836`; both pushed to `origin/main` before this handoff commit.
- Active slice: Plan 022 Slice 4ae, automated GREEN; visible picker/identity
  retest is pending.
- Rebuilt and registered Core/Operations packages are in `deploy/current`;
  Excel is closed.
- Preserved uncommitted user work: code `AGENTS.md` and the NAS/server note at
  the top of Architecture v4.11.

# Decisions and constraints

- Production Item Search opens only from a valid **Acceptable Managed Item n**
  cell: any numbered pair for INPUT/REQUIREMENT, pair 1 for OUTPUT.
- OUTPUT Name is descriptive only and never opens search.
- INPUT, REQUIREMENT, OUTPUT, and INSTRUCTION worksheet rows use one table-wide
  three-character Base-36 ID namespace. A two-pass allocator preserves existing
  valid unique IDs and repairs blanks, invalid IDs, and duplicates.
- Requirement and Output IDs also share one Process-definition namespace.
  Historical nonconforming draft IDs normalize on load; save validation remains
  strict.
- JSON persistence now quotes string variants before numeric detection, so
  `001` survives save/reopen. Numeric Qty/Percent/YieldBasis fields remain
  explicit numeric variants.

# Evidence and traceability

- Focused source RED `0/6`; packaged RED `1/2` recorded
  `OutputNamePickerSuppressed=False` and `UniqueRowIds=False`.
- Focused GREEN `6/6`; packaged Production callback/restart `2/2` records
  `OutputNamePickerSuppressed=True`, `UniqueRowIds=True`,
  `FirstAssignedIdRetained=True`, picker/SKU round-trip, and no physical key.
- Regressions: XLAM `74/74`, Ribbon/compile `142/142`, live roles `47/47`,
  ordered Release 1 chain `30/30`, launcher contracts `24/24`, dedicated NAS
  `16/16`, deterministic static `19/19`, reviewed growth `13/13`.
- Static metrics: 153 components, 5,079 procedures, 1,038 candidates; literal
  and unresolved dynamic-call ratchets did not regress.

# Do Not Repeat

- Do not restore OUTPUT Name as a convenience picker target.
- Do not allocate INPUT and OUTPUT IDs from separate dictionaries.
- Do not rely on cell text formatting alone for `001`; generic JSON numeric
  detection previously stripped leading zeroes after persistence.
- Legacy fixture IDs such as `OUT-A` and `REQ-B` are not the operator contract;
  packaged fixtures now use three-character Base-36 IDs.

# Assumptions to Re-verify

- Confirm normal Excel startup loads the registered `deploy/current` package
  hashes rather than a cached older add-in instance.
- Existing open Process tables are corrected by the current change handler;
  recreating a table should not be necessary.

# Open questions and blockers

- No automated blocker remains. Visible confirmation of the two picker cells
  and row-ID behavior is unresolved.

# Immediate next action

Open Production and confirm OUTPUT Name does nothing, OUTPUT **Acceptable
Managed Item 1** opens search, and all INPUT/OUTPUT/INSTRUCTION IDs are unique
after entering rows in a different order.

# Critical references

- `src/Production/Modules/modProductionProcessWorksheet.bas`
- `src/Production/Forms/frmProduction.frm:TestProcessWorksheetOutputPickerContract`
- `src/Production/Modules/modProductionJson.bas:JsonValue`
- `tests/tooling/Test-Plan022Slice4aeProcessPickerTargetsAndRowIds.ps1`
- `tests/integration/plan022_slice4ae_process_picker_targets_row_ids_green_results.md`
- `tools/validate_plan022_packaged_launchers.ps1:ProductionReusable`
- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md` Slice 4ae
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` Slice 4ae and section 8.2
