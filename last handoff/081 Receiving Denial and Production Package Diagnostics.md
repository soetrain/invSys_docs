# Receiving denial and Production package diagnostics

## 1. Goal and release outcome

Achieve Release 1 acceptance under Architecture v4.11 / Plan 022, including
comprehensive Operations/Admin Events and How-To/Diagnostic/Compare Action Paths.
Goal remains active. This 4be.1 denial implementation checkpoint is not a completed
slice: full reusable Production acceptance remains unresolved.

## 2. Current verified state

- Last verified 2026-09-08: code main **dec08fe**, docs main **25abfac**, pushed.
  This handoff/pointer is committed afterward. Normative clarification **f918687**
  was pushed before implementation; controls now v1.81.
- Code clean; Excel closed. Preserve unrelated docs: handoff 067 has 3 additions/
  3 deletions; critique 023 remains untracked. Neither was staged or edited.
- Candidate `deploy/validation-receiving-launcher-denial`: focused 133/133,
  native-dialog 137/137, full activity 845/845; all prior 771 and discovery 105
  check identities retained. Both actual fixed dialogs and normal form inspected.
- Original candidate builds/compiles/cold-starts; smoke 86/86, live-role 48/48,
  ordered chain 30/30, Viewer, three-size layout and launchers 3/3 pass.
  All three inspected Production layout PNGs match the prior accepted captures.
- Rebuild `deploy/validation-receiving-denial-rebuild`: build/cold-start/five
  compiles PASS; all 168 extracted source components match original candidate.
  Its full Production fails; broader replay and second run were not started.
- Both candidates' ten package hashes remain unchanged. Accepted deployment,
  NAS and operational workbooks remain unchanged. Runtime reports are ignored.

## 3. Decisions and constraints

D18 semantic inheritance remains approved, with no pending architecture approval.
Receiving's real generated callback records REQUESTED before its same Core cached
RECEIVE_POST guard, exactly once before owner work. Rejection produces existing
catalog-6 RECEIVING_OPEN / RECEIVE_OPEN_DENIED, Blocked/Unchanged, empty refs.
getEnabled, other guards, direct compatibility and pre-sign-in exclusions remain.
Interrupted completion cannot rebind context. Optional tracking never authorizes,
provisions, retries or blocks the owner decision. Native fixed failure notices are
observed through the actual callback; test instrumentation is unsaved.

All D5/D12/D13/D14, headless authority, exact System_Key, unknown-column,
captured-workbook and packaged launcher reuse invariants remain binding.
No speculative Production change or regression exemption is approved by evidence.

## 4. Evidence and traceability

[Curated evidence](../../invSys_fork/tests/integration/plan022_slice4be_receiving_launcher_denial_results.md)
contains commands, candidate hashes, exact gates and diagnostic limits.

- Meaningful unchanged-package RED: 108/19, expanded 112/21, no harness exception
  -> missing denial records, optional notice and original interrupted attempt
  -> D18 owner-fact/captured-context requirements -> focused 133/133 GREEN.
- Production failure matrix: original full candidate twice crashes at reusable
  surface/batch scale; prior navigation/identity candidate 2/2; candidate Core
  with other prior packages 2/2; candidate Operations with other prior packages
  2/2; clean same-source rebuild crashes at batch scale; unsaved source-preserving
  Operations invalidation/recompile probe also crashes at batch scale.
- All four failed native runs report RPC 0x800706BE after Excel exits and
  EXCEL.EXE / ntdll.dll / c0000028 application faults. Root cause remains unknown.
  The two substitution sets are diagnostics, not release/cold-start acceptance.
- No VBA edits occur in the standard Production harness; no runtime call enables
  MouseScroll. Operations recompilation alone did not resolve the crash.
- Static: 175 components, 5,501 procedures (+1 typed BeginOpen), 1,097 candidates,
  195 duplicate groups, 45 unresolved/eight literal dynamic targets; all 28 prior
  oversized-module caps hold. Historical Slice12 retains the same six failures
  on old/new baselines (7/13); current ratchets pass, not a claimed historical GREEN.
- Source checks: control surface 6/6, cutover 14/14, Receiving 10/10, tooling 62/62;
  PowerShell parsing and diff checks pass. No generated runtime reports committed.

## 5. Do Not Repeat

Do not repeat identical full candidate runs, a clean same-source rebuild, or
Operations-only recompilation hoping for GREEN. Preserve every failed run.
Do not treat source equality or a diagnostic substitution success as a waived
Production gate. Do not infer a native crash is a meaningful D13 behavioral RED.
The suspected mouse hook / harness edits supplied no cause. Keep existing native
dialog observer; the older UI Automation issue was already addressed separately.
Allow normal Excel cleanup between validators; never kill unidentified Excel.
Never rebuild loaded packages or promote diagnostic combinations implicitly.
Handoff 080's worksheet-input calibration failures remain independently applicable.

## 6. Assumptions to re-verify

Git/Excel state, package hashes and actual loaded dependency paths. Substitution
diagnostics preloaded Core; their result does not prove cold-start binding.
Native worksheet Confirm delivery is still unproved. Cached compilation was a
hypothesis; the Operations-only probe failed. No native debugger was discovered
on PATH. Do not collect dumps or arbitrary audit/operational values.

## 7. Open questions and blockers

Production package interaction blocks this checkpoint's completion, not progress
on the full Goal. Native worksheet activity, remaining Operations/Admin coverage,
publication, Event Tracking Settings, comprehensive Viewer, recordings/conclusions,
guides, both presentations and physical multi-station/NAS/user UAT remain open.
No Goal complete/blocked transition is warranted after this progress.

## 8. Immediate next action

Run a bounded full Production diagnostic with candidate Core and Operations
together against prior unchanged Admin/Domain packages, asserting actual loaded
dependency provenance before interpreting its result.

## 9. Critical references

- Architecture D18 launcher-denial clarification; Plan 022 4be.1-4be.6;
  controls v1.81; handoffs 079/080 for accepted baseline and worksheet limitations.
- `modReceivingActivityAction.BeginOpen`, `modTS_Received.ShowReceivingForm`,
  `modReceivingActivityCodes.LifecycleOutcome`, generated `CapabilityGuardInAction`.
- `Slice4beReceivingLauncherDenial.ps1`, `Slice4beReceivingDenialDialogs.ps1`,
  `Test-Slice4beConfigCommands.ps1` switches in curated evidence.
- Ignored `reports/runtime/slice4be-launcher-denial/`: focused red/green,
  activity/native results, `receiving-visible/`, `production/`, `production-retry/`,
  `production-baseline/`, `isolation-Core/`, `isolation-Operations/`, `rebuild/`,
  `recompile-probe/`, source/hash and static comparisons, redacted crash summaries.
- Ignored diagnostic scripts: `run-package-isolation.ps1`, `rebuild-probe.ps1`,
  `create-recompile-probe.ps1`, `recompile-production-probe.ps1`.
  `rebuild-gates.ps1` is prepared but was not run after rebuild failure.
- Diagnostic package directories use `deploy/diagnostic-launcher-denial-v2-*`;
  specific local `.git/info/exclude` entries keep these binaries untracked/ignored.
  Preserve original candidate and prior `deploy/validation-receiving-navigation-identity`.
