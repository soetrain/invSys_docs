# Goal and release outcome

Achieve invSys Release 1 acceptance under Architecture v4.11; Slice 4an now
allows one Process to combine independently measured material and packaging UOM
groups without inventing a conversion.

# Current verified state

- Code and docs are on `main`; pre-slice commits were code `90a782a` and docs
  `7a00310`. Slice 4an commits are pushed at session end.
- Active slice: Plan 022 Slice 4an, automated GREEN and awaiting visible UAT.
- Last verified 2026-08-29: Excel is closed after rebuild/validation.
- Preserve unrelated code `AGENTS.md`, the normative-spec NAS note, and
  untracked `invSys dev notes.txt` as user work.

# Decisions and constraints

- INPUT formula rows are grouped by normalized UOM. Each UOM group has its own
  basis subtotal and percentages totaling 100.0%.
- A Process may combine LB material with EA packaging and produce EA output.
  Unlike UOM quantities are never added or implicitly converted.
- Recipe connections still require UOM compatibility; a true conversion remains
  an explicit Process.
- Retrieval preserves every row quantity/UOM and deletes only the successfully
  selected table under the captured saved-workbook boundary.

# Evidence and traceability

- Focused source: `1/7` RED, then `4/7` RED after public-handler expectations;
  final `7/7` GREEN. Superseded Slice 4y remains `6/6` GREEN.
- Packaged Production: `0/2` behavioral RED with the exact one-compatible-UOM
  message; final public handler plus clean restart `2/2` GREEN.
- Fixture: 4.5 LB concentrate + 1 EA bottle + 1 EA cap -> 1 EA finished bottle;
  LB basis/percent 4.5/100%, EA basis/percents 2/50%/50%.
- Packaged XLAM `81/81`; Ribbon/compile `142/142`; live roles `47/47`;
  ordered Release 1 `30/30`; NAS `16/16`; static `19/19`; growth `13/13`.

# Do Not Repeat

- Do not restore the global `inputUoms.Count > 1` rejection.
- Do not total percentages across unlike UOM groups; a valid LB+EA Process totals
  100% within each group, not 100% across the entire table.
- Do not loosen Recipe-edge UOM compatibility as a workaround.

# Assumptions to Re-verify

- The user's existing table `invSys_Process_004_20260827` remains available in
  the saved Production operator workbook for visible retest.

# Open questions and blockers

- No automated blocker. Visible retrieval/save/release acceptance is pending.

# Immediate next action

Open Production and use **Retrieve Selected Process** on
`invSys_Process_004_20260827`; confirm its LB and EA rows import and the selected
table is removed, then save/release the Process.

# Critical references

- `src/Production/Modules/modProductionProcessWorksheet.bas`
- `src/Production/Forms/frmProduction.frm`
- `tests/tooling/Test-Plan022Slice4anMixedUomProcess.ps1`
- `tools/validate_plan022_packaged_launchers.ps1`
- Architecture v4.11 D15
- Plan 022 Slice 4an
