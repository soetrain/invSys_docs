# Slice 4be D5 commands validated

## 1. Goal and release outcome

Advance invSys Release 1 acceptance under Architecture v4.11 and Plan 022 while
preserving accepted roles and exact inventory identity. The approved D5
configuration-command prerequisite is implemented and technically validated;
the comprehensive Event Viewer/recorded Action Path contract remains pending.

## 2. Current verified state

- Last verified 2026-09-06: code `main` at `b4ce6d9`; documentation `main` has
  the synchronized D5 contract/evidence commit `7b0aee1`. This continuation and
  CURRENT pointer are committed afterward; resolve their commit with Git.
- The five XLAMs were rebuilt in `deploy/current`. Its manifest records their
  checked hashes. No operational workbook, NAS package set, or installed station
  package was intentionally modified for this slice.
- Code working tree was clean after the implementation commit. Preserve the
  user's existing 3-addition/3-deletion edit to
  `067 Partial Goal Action Path NAS Contract.md`; it was never staged.
- All owned validation Excel instances were closed. Recheck Excel before any
  subsequent build/deploy. Raw diagnostics and screenshots remain ignored in
  `reports/runtime/config-commands` in the code repository.

## 3. Decisions and constraints

- The user's approval following the D5 explanation authorizes a read-only Config
  API plus a separate authorized headless **Core** command service. Admin owns
  UI/orchestration. This replaces the earlier Admin-owned writer proposal.
- `modConfigCommands.UpdateConfigValue` checks sign-in, `ADMIN_MAINT`, selected
  target/captured warehouse-station, full required configuration, typed values,
  identity protection and writable/clean authority before persistence. The old
  `modConfig.UpdateConfigValue` is forwarding compatibility only.
- Production UOM publication retains its validated, narrowly scoped `PROD_POST`
  route and a single catalog/version save. Ordinary reads never provision,
  repair, format or save Config; explicit setup is separate.
- D5 approval does **not** approve the separate D18 replacement. Current D18's
  curated training-record contract remains normative. Do not invent a combined
  curated/captured model from the user's comparison note.
- Preserve D12/D13, captured workbooks, launcher reuse, headless dependencies,
  exact immutable `System_Key`, unknown columns and accepted Production behavior.

## 4. Evidence and traceability

See the sanitized [D5 validation record](../../invSys_fork/tests/integration/plan022_slice4be_d5_config_commands_results.md).

- D13: initial pre-implementation RED 5 failures/10 checks; expanded baseline
  comparison 9 failures/17 checks; required-schema RED 1 failure/18 checks;
  final packaged **18/18 GREEN** with actual Settings and UOM form handlers.
- Final distribution: cold-start Operations dependency check and **5/5** explicit
  project compiles; **81/81** serial packaged smoke; **48/48** live-role checks;
  **30/30** ordered full Release 1 chain with restart/reconciliation.
- Viewer passes on the final build. Layout passes at three sizes across five
  tested Production pages, including native window actions. All-role launcher
  checks pass **3/3**. Long reusable Production checks pass **2/2 aggregate
  assertions**, including persisted exact released Recipe after clean restart.
  All 150 other packaged component code hashes match that long-regression
  build; the sole final delta is the separately protected four-line required
  configuration guard in `modConfigCommands`.
- Static maintenance: 1,077 candidates, 192 duplicate groups, 45 unresolved
  dynamic calls, 8 literal Application.Run targets and 28 oversized-module
  ratchets; no growth or exception. No scanner-directed code deletion occurred.
- The actual Settings form was captured and visually inspected with BatchSize
  601 and **Configuration saved.** This is disposable-fixture operator evidence;
  fresh human UAT and overall Release 1 acceptance are not claimed.

## 5. Do Not Repeat

- Package construction is not a full compile. The explicit compiler exposed
  pre-existing Shipping diagnostic references, a Production parameter typo and
  two unqualified event constants; their minimal repairs are in `b4ce6d9`.
- Copying a candidate XLAM can retain its absolute Core reference. Preloading
  the correct Core masks that problem. Build in the intended directory and run
  the cold-start probe before preloading Core.
- Validate all required configuration before writing, not only the edited key.
  A missing WarehouseName previously allowed a different setting to save.
- A live fixture expecting LoadConfig to create a missing workbook contradicted
  approved D5. It now checks non-creating failure, then explicitly provisions.
- Keep Excel validators serial. One smoke attempt lost its COM connection and
  produced eight dependent failures; a fresh serial run passed 81/81 without
  code changes. The cause is unconfirmed; do not call that behavioral RED.

## 6. Assumptions to re-verify

- Current package hashes, Excel process state and Git status before further work.
- NAS/station deployment and real operator acceptance were not re-established
  by these isolated tests. Do not treat fixture success as new NAS UAT approval.

## 7. Open questions and blockers

- Explicit decision on the D18 replacement proposal: actual control capture,
  comprehensive R1 history and Admin-configurable Event Detail are still proposed.
- Fresh human acceptance of D5 behavior and the broader Release 1 goal remain
  open. The earlier goal was blocked on contract approval; completed D5 work
  must not be mistaken for completion of the Event Viewer slice or full goal.

## 8. Immediate next action

After explicit D18 approval, create and run the first focused packaged
Viewer/Admin form-handler test for the approved Event Detail/Action Path behavior
before implementing it.

## 9. Critical references

- `0 plan docs/xlam_invSys/CURRENT_SPEC.md` -> `invSys-Design-v4.11.md`, D5 and D18.
- `expert guidance docs/CURRENT.md` ->
  `022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`, Slice 4be.
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`, version 1.59.
- Code `src/Core/Modules/modConfigCommands.bas`, `modConfig.bas`, `modUomSettings.bas`.
- `frmAdminSettings.mBtnSaveConfig_Click`; Production UOM Send/Retrieve handlers.
- Code `tests/tooling/Test-Slice4beConfigCommands.ps1`,
  `Test-PackagedVbaCompile.ps1` and the linked sanitized validation record.
