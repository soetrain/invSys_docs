# Tscon desktop test: error5 stop

## 1. Goal and release outcome

Complete Slice4be-A under Architecture v4.11/Plan022; the next acceptance task is
General Settings independent recording and guide-conclusion proof. The user asked
to attempt that work while testing RDP disconnection after their reported
`tscon.exe 1 /dest:console`, and to exit immediately on desktop error5.

## 2. Current verified state

Last verified:2026-10-08,21:31 UTC (14:31 PDT).

- Code main `48ae1cd4`, pushed and clean; runtime checkpoint remains `6f583238`.
- Docs main prior to this handoff: `9cc3792`, pushed. Plan022/Controls1.499 unchanged.
- **Stopped before implementation or Excel tests.** Only read-only preparation and
  an ignored desktop diagnostic ran. Excel process count is zero; monitor exited.
- General647/UOM228 acceptance evidence from handoff120 remains valid at its frozen
  candidate. No new recording/guide proof was completed.
- Preserve user-owned handoff067 changes, old report024 deletion and untracked
  critique023/REPORT024/guidance025; none is staged with this handoff.

## 3. Decisions and constraints

Honor the user's error5 stop condition. No RDP/session/lock settings were changed,
no input or keep-awake activity was generated, and no screenshots were saved.
The user's tscon invocation/target was not independently inspected. This attempt
did not preserve desktop access for the agent; do not claim tscon universally fails.
D18-REPLAY-01 remains approved; separate Scale/UI decisions remain pending.

## 4. Evidence and traceability

- Last successful cursor/input-desktop probe: **21:30:40.3802982 UTC**
  (**14:30:40 PDT**), both error codes0.
- First monitored failure: **21:31:11.7284945 UTC** (**14:31:11 PDT**).
  Cursor error5; input-desktop error5; capture Win32 error6; session-state value4.
- The monitor stopped on this first sample: one sample, one error5 sample.
  This bounds observed loss of access to approximately31 seconds after the prior
  successful probe; it does not establish the precise disconnection time.
- Ignored code evidence: `reports/runtime/rdp-desktop-probe-77720af1651c4a5791c3d9fd12d26ca5/`
  (`samples.jsonl`, `summary.json`). No operational data or images were extracted.

## 5. Do Not Repeat

Do not continue interactive work after this failure, alter lock settings, keep a
probe running after exit, or infer success from tscon alone. Do not rerun unchanged
General647/UOM228 gates merely to recover context.

## 6. Assumptions to re-verify

Restore and verify desktop access before Excel interaction. Any later session
transfer/disconnection must be evaluated from fresh access/capture evidence.

## 7. Open questions and blockers

Interactive desktop access is unavailable. Whether the reported tscon command
transferred the session used by this agent is unverified. General Settings
recording/guide conclusions and broader4be-A acceptance remain open.

## 8. Immediate next action

After desktop access is restored, resume the packaged General Settings recording
and guide-comparison test work described in handoff120, retaining the error5 stop rule.

## 9. Critical references

- Handoff120 for verified General647/UOM228, constraints and exact next test scope.
- `tests/integration/plan022_slice4be_general_settings_results.md`.
- `tests/integration/plan022_slice4be_remaining_acceptance.md`.
- Architecture v4.11 D18 catalog31; current Plan022 and Controls1.499.
