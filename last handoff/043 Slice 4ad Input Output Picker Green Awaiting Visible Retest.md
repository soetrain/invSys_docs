# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Plan 022 Slice
4ad makes Production Item Search reachable from the same visible managed-item
column for INPUT and OUTPUT Process rows.

# Current verified state

- Last verified 2026-08-25: code `main` at `dd4649a`; normative/plan/control
  docs commit `4da41d9`; both are pushed to `origin/main` and this handoff is
  the current docs pointer.
- Active slice: Plan 022 Slice 4ad, automated GREEN; visible INPUT/OUTPUT picker
  retest is pending.
- Rebuilt Core and Operations packages are in `deploy/current`; Excel is closed.
- Preserved uncommitted user work: code `AGENTS.md` and the NAS/server note at
  the top of Architecture v4.11.

# Decisions and constraints

- **Acceptable Managed Item 1** is the common visible picker entry point for
  INPUT and OUTPUT rows. INPUT continues to use numbered alternatives.
- For OUTPUT, the selected item is mirrored visibly and its SKU is stored in
  hidden Accepted SKU 1 and canonical hidden Output SKU.
- A nonblank OUTPUT Name remains a descriptive definition name and is not
  overwritten when selection begins in Acceptable Managed Item 1. Name remains
  an optional convenience picker target.
- Picker selection stores no source `System_Key`; completion creates a new key.

# Evidence and traceability

- User-visible symptom: INPUT opened search; OUTPUT did not.
- Corrected source RED: `5/6`; corrected packaged public-handler RED: `0/2`
  with `OutputPickerOpened=False`, `OutputPickerCommitted=False`, and
  `OutputSkuRoundTrip=False`.
- Focused GREEN: Slice 4ad `6/6`; prior 4ac/4ab/4aa/4z remain `6/6`, `4/4`,
  `8/8`, and `7/7`.
- Packaged public Production callback plus clean restart: `2/2`, recording open,
  commit, SKU round-trip, retained descriptive Name, and no physical key true.
- Regressions: XLAM `74/74`, Ribbon/compile `142/142`, live roles `47/47`, clean
  Release 1 chain `30/30`, launchers `24/24`, NAS `16/16`, static `19/19`, and
  reviewed growth `13/13`.

# Do Not Repeat

- Do not test only OUTPUT Name; that was Slice 4ac's false-positive target and
  missed the operator's same-column workflow.
- Do not treat the OUTPUT managed item as an INPUT alternative even though its
  visible selector reuses Acceptable Managed Item 1.
- The first full-chain attempt passed all business assertions but hit the known
  test cleanup exception. Closing only the verified `/automation -Embedding`
  process and rerunning from clean state produced `30/30`.

# Assumptions to Re-verify

- Confirm installed Excel loads the rebuilt Core/Operations packages rather
  than cached earlier add-ins.
- Existing open Process tables need the current package event handler; recreating
  the table should not be required because the visible column already exists.

# Open questions and blockers

- No automated blocker remains. Visible confirmation that search opens from
  both INPUT and OUTPUT rows is unresolved.

# Immediate next action

Open Production, enter **Acceptable Managed Item 1** on an OUTPUT row, select an
item, and confirm **Retrieve Selected Process** succeeds with its Name retained.

# Critical references

- `expert guidance docs/022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md` Slice 4ad
- `0 plan docs/xlam_invSys/invSys-Design-v4.11.md` D15 worksheet contract
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md` Slice 4ad and section 8.2
- `src/Production/Modules/modProductionProcessWorksheet.bas`
- `src/Core/ClassModules/cDynItemSearch.cls`
- `src/Production/Forms/frmProduction.frm`
- `tests/tooling/Test-Plan022Slice4adProcessPickerRecordTypes.ps1`
- `mProduction.RunProcessWorksheetOutputPickerContractTest`
