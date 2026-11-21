# Action Paths Search: native performance failure after packaged GREEN

Date: 2026-11-20 Pacific. Status: **initial slow report; candidate prompt in native checks and qualitative 29-recording NAS review; recording-row usability open**. This is a diagnostic record, not a proposed contract change.

## 2026-11-21 visible NAS review

Justin reports Search recordings working with the 29-recording disposable
WHT7025AE library. A screenshot shows a query filtering several rows. This is
qualitative operator evidence, not a measured latency result or final 4be-A
acceptance. The visible row fields are the fixed **Recorded sequence** title
and UTC creation time; similarly named recordings cannot be identified by
task from the list. The guide picker and guide editor provide separate
authoring paths and do not improve these saved-recording rows.

## 2026-11-20 native keyboard result

The test-only fixture now saves disposable instrumented packages before Viewer
entry and keeps a saved operator host open. A process-owned click focuses the
Search textbox once; the six subsequent keys are sent as a continuous native
typing sequence. The actual packaged form and retained Viewer/recording-reader
checks passed **152/152**. Seven disposable recording headers were loaded.
First key after selection, two later letters, Backspace, Clear and a nonmatch
each had exact textbox/list/selection results. Key-to-result times were 48.6,
6.1, 5.6, 5.7, 49.7 and 4.5 ms, respectively. All five instrumented XLAMs
compiled. Excel exited naturally, the saved operator host's identity and bytes
were preserved, and the five candidate XLAM hashes were unchanged. Ignored
local evidence: `reports/runtime/slice4be-viewer-published-read/717cc49005144623916f65c9630c3867/`;
the timing JSON SHA-256 is `21EBBDCF7E6EAE857DA7AE373C734043A177AAC5951CE71616E0494467EC9AC1`.

Earlier fixture failures were not behavioral RED: unsaved instrumented packages
failed at Viewer caption assignment; no native focus left the first text empty;
repeated textbox clicks replaced a later character; an assumed one-row result
for `Rec` was wrong because all disposable rows were titled Recorded sequence.
The correction changes only test tooling. This result does not prove the
historical slow session's loaded package identity, qualify a larger NAS library,
or close 4be-A acceptance. `deploy/current` and the NAS were not changed.

## 2026-11-21 larger-library NAS setup

In disposable WHT7025AE, guarded clicks on the real Viewer Start and Stop
buttons created 28 additional empty recordings. The library now contains 29
distinct recordings; all have Stopped Close entries and none remains active.
The initial one-recording pilot and 26 of the next 27 runs closed within their
bounded sequences. The final run's first Stop click did not take effect; a
separate guarded Stop closed it after inspection. Both fill reports confirmed
pre-existing file hashes were preserved. The ignored local setup reports are
under `reports/runtime/action-search-scale-20261120/` in the code worker.
This is a search-load fixture, not a How-To guide or a native Search result at
29 rows. Visible operator responsiveness remains pending. No inventory action,
product source, XLAM or approved contract was changed.

## Operator observation

On the disposable NAS review warehouse, **Action Paths > Search recordings** remains extremely slow to type into after the D18-SEARCH-INPUT-01 change. Justin reports that **Viewer > Events > Search** is fast and that the previously built picker is also responsive. No per-character timing, Excel CPU trace, or exact loaded-add-in inventory was captured during this observation. Do not label the final build visibly accepted.

## Contract and candidate

Architecture v4.11, approved **D18-SEARCH-INPUT-01**, requires immediate local filtering of recording headers loaded by authorized Open/Refresh. Each edit checks captured session/warehouse and clears selected evidence/actions, but makes no per-key Config or NAS journal read. Selection and Refresh recheck current policy. Plan 022 and Controls 1.785 recorded packaged GREEN with visible NAS usability still pending; the initial operator result made that visible gate **RED at the time**. The later candidate retest is recorded below.

Code commit `aa44b245` is on draft [code PR17](https://github.com/soetrain/invSys_fork/pull/17). The isolated validation folder is `C:\Users\Justin\repos\invSys_edit_receipts\deploy\validation-action-search-final-20261120\`. Its `invSys.Operations.xlam` SHA-256 is `7A174635017618DDCC54B6F03D831780FA16B3F9D117F05E27962D97C9B3266D`; `deploy/current` has a different Operations hash, `882DABBDFF327C1B280CD4D13428577969B89ECBC576AE89C6263FC117ABB9C8`. **The exact Operations XLAM loaded in Justin's slow session has not been independently verified.**

## What the existing evidence proves

- Prior-package actual-handler RED: **123 PASS / 4 behavioral FAIL**. The old `mSearch_Change` called Core `ListPaths` on each edit, validating policy and rereading journal headers.
- New isolated candidate: **130/130 packaged GREEN**, five explicit XLAM compiles, cold-start reference check, layout and integrity regressions, natural Excel exit. The test sets `txtPathSearch.Value` through the packaged form adapter and exercises its real `Change` handler. It verifies no per-key policy read and that filtering still works when the disposable journal files are temporarily absent. Evidence: ignored local `reports/runtime/slice4be-viewer-published-read/f3256936b0ec44b9a7e4bfe935c61bf2/` in the worker.
- These checks **do not measure native keyboard delivery, focus, repaint, RDP effects, or key-to-visible-result latency**. They do not establish that the operator loaded the new package. Static maintenance did not show dynamic-call growth: 391 components, 6762 procedures, 9 literal/45 unresolved dynamic calls, 193 duplicate-body groups.

## Source path to examine

In `src/Operations/Forms/frmActionPaths.frm`, `mSearch_Change` (near line 276) calls `ContextValid`, `ClearSelectedRecording`, then `FilterLoadedPaths`. `FilterLoadedPaths` (near line 132) clears the MSForms ListBox and repopulates matching rows from `mLoadedRows` on each edit. `mPaths_Change` is suppressed while `mLoading` is true. `ClearSelectedRecording` can close dependent forms and clear selected evidence on the first edit after a selection. `modActivity.CaptureContext` uses the in-memory sign-in/target state and path-string comparison; the inspected path has no direct Config or journal read. A hidden or indirect cost has **not** been ruled out by native timing.

Comparators in the same Operations XLAM: `frmInventoryViewer.mTxtSearch_Change` calls local `RenderRows` and is reported fast. `frmGuideActionPicker.mSearch_Change` calls `RefreshChoices` and is also reported fast. Their existence argues against assuming that all MSForms list rebuilding, Core calls, or NAS use is intrinsically slow. No causal conclusion follows without timings in the same operator session.

## Questions for expert review

1. Confirm the five loaded XLAM full paths and hashes inside the slow Excel session. Is `invSys.Operations.xlam` the isolated `7A174635…` candidate, with no older registered copy supplying `frmActionPaths`?
2. Measure one native keypress from input delivery to visible list update, and separately time `ContextValid`, `ClearSelectedRecording`, `FilterLoadedPaths` (including ListBox Clear/AddItem and repaint), and any event reentry. Compare first key after a selection with later keys. Capture counts and timings only; do not log typed text or operational rows.
3. Repeat comparable keypresses in Action Paths Search, Viewer Events Search, and the working picker under the same Excel/session/display state. Record loaded row counts and whether RDP or console is active.
4. Identify the first measured slow boundary, then propose a **focused native-keyboard RED** through the packaged form and a correction that preserves D18-SEARCH-INPUT-01's session, policy, evidence-clearing and Refresh rules. Re-run native NAS usability and retained packaged regressions before accepting a fix.

Do not infer a new search contract or promote `deploy/current` from the 130/130 packaged result. The underlying slow boundary is unresolved.

## 2026-11-20 startup-registration follow-up

Read-only inspection found Excel startup slots `OPEN` and `OPEN1` pointing to the November 16 review copies of Admin and Operations, while `OPEN2`–`OPEN4` pointed to the November 20 Core and Domain candidate. The registered Operations file hashes to `8B634BE75820C186A8A42EF434CDCFC681BC8857086A56A3DE424ECC02D348AC`; its local candidate manifest names source commit `45d84ca8`. At that commit, `frmActionPaths.mSearch_Change` calls `RefreshPaths` on each edit, which calls Core `ListPaths`. The registered November 20 Operations file instead hashes to `7A174635017618DDCC54B6F03D831780FA16B3F9D117F05E27962D97C9B3266D` and filters the loaded snapshot locally. This is a concrete, plausible explanation for the slow review, **not proof of the exact workbook loaded in Justin's earlier Excel session**.

With Excel closed, the two old startup entries were saved to ignored local `reports/runtime/action-search-startup-before-20261120.json` in the worker and temporarily changed to the November 20 candidate paths. All five startup slots now name that same candidate folder. `deploy/current`, the NAS and package files were untouched. The next manual Excel session must verify the actual five loaded workbook paths/hashes before a native Search conclusion. Restore the saved entries when the review no longer needs the candidate.

A new test-only assertion that programmatic Search leaves no row selected passed in the existing hidden-host reader suite: **136/136**, natural Excel closure. It rules out automatic selection in that fixture. A separate native-keyboard probe is staged behind an explicit test switch but has **not reached Search**. Three visible-host attempts stopped in `modInventoryViewer.OpenInventoryViewer` before the probe with a captured Visual Basic `80010007` automation dialog; one attempt also stopped earlier on a test-only probe insertion error that was removed. The visible failures persisted after the startup entries were corrected. Their cause is unconfirmed and they are not behavioral Search RED. No product source was changed.

The subsequent manual Excel session's Add-ins collection reported all five November 20 candidate paths with `Installed=True`; their file sizes matched the pre-launch candidate and the hashes were captured before Excel locked the files. On disposable WHT7025AE with one saved recording, Justin reported normal, prompt typing, a nonmatching query clearing the list, prompt Backspace/clear restoration, and prompt selected-evidence/action clearing without an error. This is visible small-library usability evidence for the verified candidate, not a measured native latency or larger-library qualification. The earlier slow session's loaded path remains unknown.

Current disposition: the earlier visible slow result remains a valid observation, while the verified candidate now passes the limited operator interaction. Do not reproduce the old per-key Core path by restoring mixed startup add-ins during review. The separate `80010007` fixture failure is localized below; correct its setup before measuring native input on a larger recording library or calling the native D13 probe GREEN.

## 2026-11-20 Viewer startup stage diagnosis

The existing `-TraceViewerStartupForTest` probe on the unchanged candidate identified the failing statement inside `frmInventoryViewer.SetWarehouse`: `Me.Caption = "Viewer - " & mWarehouseId`. Initialization, generation assignment, context clearing/capture and recording binding completed before the failure; `RefreshInventory` and `Show` were not reached. The probe reports fixed stage IDs and error numbers, not workbook values.

Five focused diagnostic runs compared package and host state. Original read-only instrumented XLAMs failed at `Warehouse.Caption` with no ordinary workbook (**14 PASS / 2 diagnostic FAIL**) and with a saved active operator workbook (**16/2**). Writable disposable copies left unsaved also failed there with a saved workbook (**17/2**). Saved disposable probe copies opened Viewer successfully without an ordinary workbook (**17/17**) and with a saved active workbook (**19/19**). The active workbook was therefore not sufficient to cure the failure; saving the instrumented disposable packages before form entry is the decisive tested difference. This is a fixture/package-state finding, not behavioral Search RED or a product contract change.

Ignored local evidence is under `reports/runtime/slice4be-viewer-published-read/` in the code worker, in run folders beginning `5f6b2323`, `739e565d`, `85d6a0a7`, `a8935462`, and `254b2834`, respectively. Excel was absent after each run without assisted teardown; the five candidate XLAM hashes were unchanged afterward. The next test-only change should use saved disposable instrumented packages for the native Search fixture, then verify its actual handler and native timing. Larger-library NAS acceptance remains open.
