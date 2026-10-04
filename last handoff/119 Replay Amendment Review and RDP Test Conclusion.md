# Replay amendment review and RDP test conclusion

## 1. Goal and release outcome

Complete Slice 4be/Release 1 under Architecture v4.11 and Plan 022. Immediate next
deliverable is approval of the concrete executable How-To amendment before any
replay implementation. The desktop/monitor test is concluded at the user's request.

## 2. Current verified state

Last verified: 2026-10-04, 11:30 AM PDT.

- Code main `50241047`, pushed, clean; no runtime or test changes in this session.
- Docs main `7fac5a8`, pushed before this handoff: pending D18-REPLAY-01 in
  Architecture, matching Plan 022 sequence and proposed Controls 1.436 surfaces.
- No Excel tests/builds or operational workbook actions were started. The desktop
  probe is stopped. Check Excel state before future packaging.
- Preserve user-owned changes: modified handoff067, untracked critique023,
  user-renamed report024 (old tracked path deleted; new REPORT filename untracked),
  and untracked guidance025. They were not staged by the agent.

## 3. Decisions and constraints

User confirms the physical monitor can be off without error 5. Their reported
working arrangement for this setup is RDP connected with the session unlocked;
locking the connected computer or disconnecting RDP triggers desktop access failure.
Do not assume the physical monitor must remain on. Do not infer this is a universal
Windows rule or proof of which policy causes the session transition.

Keep the standing condition: actual desktop Win32 error 5 -> record time, stop
interactive work, safely restore owned fixtures, save/push an ending handoff and
exit. No Windows lock settings were changed. No keep-awake/input simulation was
used. The goal last reported usageLimited; no automatic goal resume is claimed.

D18-REPLAY-01 is **drafted, NOT APPROVED**. An approval question is pending; the
user's monitor conclusion does not answer it. Current approved D18 still prohibits
replay. Proposed outcome: Record -> Author -> Run in a designated training runtime
-> Verify fresh owner results, with explicit inputs/prompts, versioned profiles,
ordinary-handler dispatch, and no automatic retries/rollback. The proposal lists
the rules to replace and work to retain. Do not implement a hybrid or renumber D19.

## 4. Evidence and traceability

Monitor-off test: **452 samples**, zero cursor, input-desktop or capture failures,
and zero error 5 samples. Observed interval: **2026-10-04 18:06:19.2004883 through
18:28:58.2743182 UTC**, or **11:06:19 through 11:28:58 AM PDT**, about22m39s.
The user confirms monitor-off/RDP-connected conditions. The probe was stopped after
their conclusion; do not claim the planned full30 minutes or coverage after the
last sample. No screenshots were saved.

Ignored code receipt:
`reports/runtime/rdp-desktop-probe-d760aa0490a940b88986bbc913ab4b4b/checkpoint-summary.json`.
The preceding disconnected-session test returned cursor/input-desktop error5 on
its first sample at10:53:42 AM PDT; see handoff118 for exact evidence.

Proposal checks: matching NOT APPROVED status in all three documents; operative
no-replay clause retained pending approval; existing D19 retained; diff checks and
unrelated-file hashes verified. D13 runtime RED/GREEN does not apply to drafting an
unapproved proposal; its first packaged execution RED is specified, not performed.

## 5. Do Not Repeat

Do not keep testing the monitor after this concluded observation window. Do not
repeat unchanged accepted Excel gates or use passive diagnostics as a substitute
for the user's intended executable How-To. Do not treat recorder data as containing
inputs it never captured. Do not revisit earlier granted connection-write/UOM approvals.

## 6. Assumptions to Re-verify

Desktop access before interactive work, current RDP/session state, Excel ownership,
goal usage status, repository pointers/status and any newly supplied approval.

## 7. Open questions and blockers

Await approval or revision of D18-REPLAY-01. RUN-SCALE-01/RUN-UI-01 remain separately
unapproved. Prior Complete Run258, Check In629, smoke86, chain32/live-role48/warehouse15
and scoped layout/static evidence remain valid at their frozen candidate; human/NAS
acceptance and remaining observation/worksheet/partial-submission work are incomplete.

## 8. Immediate next action

Obtain the user's disposition on the committed D18-REPLAY-01 draft, then explicitly
update the normative conflicting clauses and synchronized records before the first
packaged replay test/implementation; recheck desktop access before interaction.

## 9. Critical references

- Architecture v4.11: pending D18-REPLAY-01 begins at line225 in commit7fac5a8.
- Plan022: pending amendment/milestone section at its beginning.
- `0 plan docs/xlam_invSys/invSys-Controls-v1.md`: version1.436 proposed surfaces.
- `expert guidance docs/025 GUIDANCE Action Path contract vs expectation conflict.md`
- `tests/integration/plan022_slice4be_production_complete_results.md`
- Handoff118 for the prior immediate desktop failure and preserved runtime baseline.
