# Comparison progress and passive shutdown crash

## 1. Goal and release outcome

Complete Release 1 and Slice 4be under Architecture v4.11 / Plan 022. The goal
remains active and incomplete. The six-workstream acceptance checklist is the
current index, not a new contract or percentage estimate.

## 2. Current verified state

- Last verified 2026-09-23 22:27 UTC: code `main` **78bbf42**, docs `main`
  **6956d0e**, pushed before this checkpoint. Controls v1.226. Runtime e986b51;
  frozen `deploy/validation-settings-diagnostic`, 299 runtime/five package pins.
- Focused diagnostics 780/780 and 54 images retain a shutdown limitation.
  Settings 202/202, owner 460/460 and UOM 228/228 have normal closure. Broader
  Boxing retains 1,707 PASS/seven pending D8-A FAIL, 22 images; host cleanup means
  its behavioral evidence is not clean-shutdown acceptance.
- Focused reader correction 124/124 retains all 116 prior checks, adding eight
  exact policy-save boundary checks; 74/74 guard calibration and normal delayed
  closure. No runtime changes.
- Comparison follow-up: **333 PASS/one harness exception**, 96 prior checks
  unreached. All reached prior checks pass. All 15 library captures and Pending/
  Partial paired views pass; 23 images reviewed. Applied pair-label read fails
  with 0x800AC472. Normal delayed closure, zero audited Application errors,
  299 runtime/192 test/five package hashes preserved. Comparison remains open.
- Passive chain is **terminal** (session53609 exit0): **32/48/15** chain/live-role/
  Create Warehouse passes, prior identities retained. Settings and three tracked
  reports restored; 299 runtime/250 tooling/five package hashes unchanged.
- Clean shutdown **fails**: original restart Excel crashes in combase.dll with
  c0000005, followed by OFFICE_MODULE_VERSION_MISMATCH. No external post-Quit COM
  reattachment or additional Quit occurs. Both events remain acceptance failures.
- User confirms Excel fully closed; native check agrees. No active test host
  requires preservation. The guarded cleanup receipt has **Targets=[]**: no
  process was terminated. A commentary claim of cleanup was corrected promptly.
- Preserve unrelated `src/Operations/Forms/frmEventDetail.frm` (+18 measurement
  lines, Locked=True), docs handoff067 modification and untracked critique023.

## 3. Decisions and constraints

- D18 semantic inheritance is approved. New controls need exact normative/Plan/
  catalog entries and D13; architectural contradictions need explicit approval.
- D8-A ordinary Auth reads/provisioning and Event Detail Locked=False amendments
  remain unapproved. Full-access/resume messages do not approve them. Carrier
  authority and exact guide-transfer wire/provenance remain unresolved.
- Preserve exact System_Key, unknown columns, captured workbooks, headless owners,
  packaged launcher reuse and all prior GREEN identities. One Excel gate at a time.
- User explicitly grants standing authorization: **"you can always close Excel."**
  Do not repeat ownership/closure permission questions. Preserve unsaved work;
  this authorization does not turn assisted test closure into clean-shutdown proof.

## 4. Evidence and traceability

- Detailed record: `tests/integration/plan022_slice4be_settings_editor_activity_results.md`.
- Comparison receipt `settings-diagnostic-fixed-regression-comparison-attempt-verification.json`;
  root `slice4be-viewer-published-read/1f026328b51f49e081c147589d4c6bca`.
- Failed getter: `frmActionPathView/lblActionPathPair/Label`, HRESULT 0x800AC472.
  Capture-helper correction passes all 15 library captures. Bounded State-probe
  calibration 31 PASS/five expected FAIL -> 36/36; no actual State retry occurs
  in the broad run. Harness RED/GREEN is not product behavioral RED.
- Static `comparison-harness-static-verification.json`: 250 components/6,045
  procedures/132,893 lines, 9/45 dynamic calls, 193 duplicate groups; three schemas,
  28 size limits, 192 PowerShell parses pass.
- Passive chain: `settings-diagnostic-chain-passive-behavior-verification.json`,
  `settings-diagnostic-chain-passive-event-facts.json`, lifecycle JSONL and empty
  cleanup receipt. Strict verifier rejects its two Application events.
- This host's Get-WinEvent query misses relevant events with UTC DateTime bounds.
  Use explicit LocalDateTime bounds. Corrected re-audit of eleven earlier timed
  receipts confirms prior counts: two in each old chain retry/trace, zero in nine
  others. See `settings-diagnostic-corrected-local-time-event-audit.json`.

## 5. Do Not Repeat

- No unchanged broad comparison or chain rerun. Isolate the remaining getter or
  restart shutdown failure first; do not automatically expand command retries.
- No COM reattachment/additional Quit after original stage Quit. Their absence
  did not prevent the crash, so earlier correlation is not a proved root cause.
- Never treat an empty cleanup target list as termination or exit0 as clean
  shutdown. The generic cleanup waits on any Excel process, including unknown
  ones. The user now authorizes Excel closure; preserve work and record any
  assistance accurately rather than treating it as unassisted test shutdown.
- Do not delete the previously rejected temporary directory ending
  `invsys-config-command-c365fc6fbfab48a79b8f842127ebf3fd`.

## 6. Assumptions to Re-verify

Recheck Excel absence, candidate pins and approvals. Desktop error5 cause remains
unknown; later successful access does not establish Codex throttling as its cause.
All recorded process IDs are historical; do not use them for future cleanup.

## 7. Open questions and blockers

Comparison COM transport, restart shutdown crash, pending contract amendments,
comprehensive Production/Admin coverage, guide transfer, remaining candidate
regressions and human/NAS acceptance remain open. See the acceptance checklist.
The Excel-ownership question no longer blocks execution: Excel is closed, and
the user grants standing permission to close it in future.

## 8. Immediate next action

Prepare a bounded restart-shutdown control with native process/event evidence
and unchanged packages to distinguish minimal Excel/package lifetime from the
restart fixture, retaining the original crash and all behavioral GREENs.

## 9. Critical references

- Architecture `0 plan docs/xlam_invSys/CURRENT_SPEC.md`; Plan
  `expert guidance docs/CURRENT.md`; controls `0 plan docs/xlam_invSys/invSys-Controls-v1.md`.
- `plan022_slice4be_remaining_acceptance.md` and the detailed Settings result record.
- Prepared unused remaining gates: `run-settings-diagnostic-comparison-fixed.ps1`
  supports curation/restart/viewer/guide; `verify-settings-diagnostic-remaining-guides.ps1`
  expects 77/27/98/355 checks and 18/3/3/52 images. Use unique unused prefixes.
- Source-only future notes: `admin-settings-remaining-coverage-notes.md`,
  `production-observation-owner-discovery.md`, `production-observation-surface-discovery.json`.
  Production projected status alone is not proof that the exact new event applied.
