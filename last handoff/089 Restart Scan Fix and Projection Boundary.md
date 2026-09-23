# Restart scan fix and projection boundary

## 1. Goal and release outcome

Complete Release 1 / Slice 4be under Architecture v4.11 and Plan 022. Goal active,
incomplete. The existing six-workstream checklist governs remaining acceptance;
the verified changes here affect test tooling, not the runtime contract.

## 2. Current verified state

- Last verified 2026-09-23 23:23 UTC: code `main` **fa151d0**, pushed; preceding
  docs checkpoint `main` **12d6507**, pushed. This handoff accompanies the next
  documentation checkpoint and controls v1.228; use Git for its final commit.
- Runtime checkpoint e986b51; frozen `deploy/validation-settings-diagnostic`.
  All 299 runtime/five package pins preserved. No new build or product compile.
- Restart verification helper fix committed/pushed in **5b18956**: explicit
  acquired COM-reference release, name-based header scan, same NoRowHeaders rule.
  Focused restart **12/12**, compatibility **8/8** after **6 PASS/two expected FAIL**.
  Both final focused controls close normally with zero Application failures.
- Subsequent real full chain fails **before the changed restart stage**: chain
  **5 PASS/one exception**, live roles **32 PASS/one exception**, Create Warehouse
  **15/15**; 27/16 preceding checks unreached. Projection deletion passes; packaged
  `modProcessor.RunBatchReportForAutomation` fails with 0x800706BE. Original Excel
  crashes in ntdll.dll/c0000028 with a subsequent Office mismatch event.
- New copied-fixture processor control **6/6 behavior**, but strict lifecycle
  **FAIL**: Excel remains after original Quit for the 210-second observation,
  then exits after worker release without assistance. Delayed audit confirms
  zero Application failures. This fresh session omits earlier role actions and
  does not resolve the full-chain crash.
- All test sessions terminal, Excel closed, local settings and three tracked
  full-chain reports restored. No test host or in-memory restoration state pending.
- Unrelated changes preserved: frmEventDetail.frm (+18 measurement lines,
  Locked=True), docs handoff067 modification, untracked critique023. Do not stage.

## 3. Decisions and constraints

- User grants standing Excel-closure authorization: **"you can always close Excel."**
  Preserve unsaved work and distinguish assisted closure from acceptance evidence.
- D18 semantic inheritance approved; discovered controls need exact spec/Plan/
  catalog entries and D13. Contradicting architecture requires explicit approval.
- D8-A ordinary Auth reads/provisioning and Event Detail Locked=False amendments
  remain unapproved. Full access/resume/closure permission does not approve them.
  Carrier authority and exact guide-transfer wire/provenance remain unresolved.
- Preserve exact System_Key, unknown columns, captured workbook binding, headless
  authority, packaged reuse, and every prior GREEN identity. One Excel gate at a
  time. No rebuild/deployment while relevant workbooks/add-ins are open.

## 4. Evidence and traceability

- Full detailed new evidence: `tests/integration/plan022_slice4be_shutdown_header_results.md`.
  It preserves every failed control/calibration and exact ignored receipt roots.
- `shutdown-header-evidence-verification.json`: 299 runtime/252 tooling/five package
  pins; 250 components, 6,045 procedures, 132,893 lines, 9 literal/45 unresolved
  dynamic calls, 193 duplicate groups; three schemas, 28 size limits, 252 PS parses.
- Failed broad prefix: `reports/runtime/settings-diagnostic-header-release-chain`.
  Ten saved generated test workbooks in a distinct recovery Excel are closed
  without saving, then its first Quit. Original crashed process untouched.
  Locked files prevented pre-closure hashes: byte preservation across closure
  is unproven, **not evidence of a changed file**. No operational workbook involved.
- Projection root `slice4be-shutdown-control/6248968ddfac470594907614703db211`:
  missing projections, rebuilt tables, one applied/log append, exact event/key,
  processed status and duplicate-free replay all pass. Three preceding fixture
  calibrations expected missing tables, but recovered files contain both tables
  and a pending trigger. They are fixture failures, not product RED.
- `projection-replay-source-verification.json`: all runtime pins unchanged; one
  intended diagnostic-script change among 252 tooling files, all parse. Delayed
  audit `projection-replay-delayed-event-audit.json` at 23:21:50 UTC: zero events,
  Excel closed. Package/settings restoration also verified in the root result.
- Governing rule remains D13 plus Release 1 projection/restart acceptance. The
  full-chain processor crash is unresolved; fresh-session success narrows context
  but identifies neither its cause nor a runtime fix.

## 5. Do Not Repeat

- No unchanged broad chain/comparison retry. The next projection diagnostic must
  preserve the exact preceding role handlers in the same Excel session.
- Indexed header enumeration alone did not fix shutdown. Header-cell reads broke
  hidden-header compatibility; retain ListColumns.Name semantics and acquired
  reference release. Do not replace the name-based helper with the rejected variant.
- Do not reattach COM or send an additional Quit to an original process during
  its passive observation. Afterward preserve any failure and record assistance.
- Use explicit **LocalDateTime** bounds for this host's Get-WinEvent queries.
  UTC DateTime queries missed real events; process exit/exit0 is insufficient.
- Do not equate worker-release-dependent closure with closure while worker alive.
- Do not delete the previously rejected temporary directory ending
  `invsys-config-command-c365fc6fbfab48a79b8f842127ebf3fd`.

## 6. Assumptions to Re-verify

Recheck Excel absence, candidate hashes, approvals and fixture availability.
All recorded process IDs are historical. Desktop error5 cause remains unresolved;
successful later access does not prove throttling caused it.

## 7. Open questions and blockers

Comprehensive Production/Admin coverage, remaining Settings/lifecycle controls,
guide transfer, pending amendments, comparison COM failure, full-chain crash,
closure limitations, current-candidate regressions and human/NAS acceptance remain.
Preserved candidate evidence: Settings diagnostics780/780 (54 images, closure
limit); Settings202/202, owner460/460, UOM228/228 normal closure; Boxing1707PASS/
seven pending D8-A FAIL (22 images, cleanup limit); reader124/124 plus74/74 guard;
comparison333PASS/one exception,96 unreached (23 images), Applied pair-label
0x800AC472 unresolved. Prior passive chain32/48/15 behavior and failed combase.dll
shutdown remain preserved. None are silently replaced by the later shorter run.

## 8. Immediate next action

Prepare a focused phase-cut reproduction of projection processing after the exact
preceding packaged role handlers in one Excel session, retaining all reached check
identities and native process/event evidence before changing implementation.

## 9. Critical references

- Spec/Plan/controls pointers unchanged; controls v1.228.
- `tests/integration/plan022_slice4be_remaining_acceptance.md` and
  `plan022_slice4be_settings_editor_activity_results.md`.
- `tools/validate_release1_full_chain.ps1`: Test-NoRowHeaders,
  New-OrderedLiveValidator, Invoke-RestartReconciliation.
- `tools/validate_phase6_live_role_workflows.ps1`: projection recovery block,
  Restore-LiveRuntimeContext; `tests/tooling/Test-Slice4beShutdownControl.ps1`.
- Prepared unused candidate gates: `reports/runtime/run-settings-diagnostic-comparison-fixed.ps1`
  curation/restart/viewer/guide; verifier expects77/27/98/355 and18/3/3/52 images.
  Use unique prefixes and resolve focused failures before broad retries.
