# Controlled cursor handoff comparison

## 1. Goal and release outcome

Resume Slice4be-A General Settings recording/guide proof after resolving native
input access; Architecture v4.11 and Plan022 remain unchanged.

## 2. Current verified state

Last verified:2026-10-08,18:17:39 PDT. Code main `f5aa178c` is pushed;
docs main before this update is `559cd2e`. Runtime remains `6f583238`.
General647/UOM228 retain their frozen-candidate scope. No Excel was started.
The diagnostic helper completed and exited normally; no helper remains running.
Preserve the user-owned documentation changes listed in handoff122.

## 3. Decisions and constraints

The user requested a before/after cursor comparison around their tscon script,
with part two30 seconds after their instruction. Both measurements ran in the
same helper process, session and native thread, using C#/.NET P/Invoke to
user32 SetCursorPos/GetCursorPos. No desktop rebinding, helper restart, Windows
settings change or Excel retry occurred. Retain the immediate error5 stop rule.

## 4. Evidence and traceability

- Before,17:46:01 PDT: SetCursorPos TRUE, exact8-pixel movement, TRUE restoration
  and verified original coordinates; display2880x1800.
- After,18:17:39 PDT: SetCursorPos FALSE, immediately captured Win32 error0,
  unchanged coordinates after immediate/50ms reads; requested target was8 pixels
  away and inside the current1024x768 display. Restore also returned FALSE/error0;
  coordinates stayed original because the cursor had never moved.
- Both phases: calling desktop confirmed as active input desktop via UOI_IO;
  names match, temporary physical-pixel DPI context restored, no error5.
- Diagnosis: cursor API failure appears after handoff despite a valid current
  target and matching active input desktop. This is not API success with unchanged
  coordinates. The display changed; its causal role is unproven.
- Exact process/thread/session IDs, coordinates, bounds and API returns remain
  in ignored `reports/runtime/cursor-handoff/15c91091bc3244919fb15eca1070dae1/`
  (`identity.json`, `before.json`, `after.json`, `stopped.txt`). No screenshots,
  credentials, operational data, clicks or keystrokes were collected/generated.

## 5. Do Not Repeat

Do not infer working input from desktop reads or absence of error5. The report's
`OriginalPositionRestored=true` is coordinate equality, not successful restoration
API delivery; read it alongside `SetCursorPosRestore.ReturnValue=false` after handoff.

## 6. Assumptions to re-verify

The user reports using tscon; its invocation was not independently instrumented.
Do not infer the underlying Windows/driver cause from this controlled comparison.

## 7. Open questions and blockers

Native pointer movement after handoff remains blocked. General recorded/guide
cases are still unreached as described in handoff122; no new product acceptance.

## 8. Immediate next action

Use the preserved comparison to select the next focused environment diagnostic;
report findings before any desktop-binding change, helper restart or Excel retry.

## 9. Critical references

- `tools/diagnose-cursor-handoff.ps1` at `f5aa178c`; helper no longer running.
- Handoff122 and `tests/integration/plan022_slice4be_general_settings_results.md`.
- Current Architecture v4.11, Plan022 and Controls1.500.
