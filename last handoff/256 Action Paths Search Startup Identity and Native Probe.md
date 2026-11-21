# Handoff 256: Action Paths Search package identity and native probe

Date: 2026-11-20 Pacific. Predecessor: handoff 255.

## Goal and release outcome

Complete 4be-A under Architecture v4.11; D18-SEARCH-INPUT-01 has a prompt small-library NAS retest, with measured and larger-library Search qualification and final acceptance open.

## Current verified state

Last verified 2026-11-20: code worker `C:\Users\Justin\repos\invSys_edit_receipts`, `work/d18-picker-save-repro` at `e4220745`, pushed to draft [code PR17](https://github.com/soetrain/invSys_fork/pull/17); only pre-existing untracked `tests/tooling/reports/` remains. Docs `work/perplexity-handoff218`, draft [docs PR8](https://github.com/soetrain/invSys_docs/pull/8); preserve unrelated guidance 024 deletion, handoff 067 edit and untracked 023/024/025/WSAIFabricSvc files. Excel is closed after the focused startup diagnostics. No product source, XLAM, `deploy/current` or NAS file changed.

## Decisions and constraints

Approved D18-SEARCH-INPUT-01 is unchanged. The user-supplied `expert guidance docs/027 Action Paths Search Problem Response.md` is investigative advice, not a new contract. Excel's `OPEN`/`OPEN1` startup entries were temporarily changed from November 16 Admin/Operations copies to the November 20 validation copies with Excel closed; `OPEN2`–`OPEN4` already pointed there. Prior values are in ignored local `reports/runtime/action-search-startup-before-20261120.json` in the worker. Restore those two entries when the review no longer needs the candidate. Do not infer the earlier live package from startup entries alone.

## Evidence and traceability

The former registered Operations XLAM hashes to `8B634BE7…02D348AC`; its manifest names source `45d84ca8`, where `frmActionPaths.mSearch_Change` calls `RefreshPaths` per key. The November 20 candidate hashes to `7A174635…9B3266D` and filters the loaded snapshot locally. This is a strong package-mismatch hypothesis for Justin's slow review; the exact file loaded then was not captured. The old and candidate Admin hashes also differ. See updated report 026.

The hidden packaged reader suite passed **136/136**, including a new no-auto-selected-row assertion, with natural Excel closure. The test-only native keyboard probe is committed behind `-CheckRecordingNativeSearch`; it has **not reached Search**. The earlier visible-host attempts stopped at `modInventoryViewer.OpenInventoryViewer` with captured `80010007` dialogs. The existing startup stage probe now localizes this to `frmInventoryViewer.SetWarehouse` at `Me.Caption`: original read-only instrumented packages failed with and without a saved host (14/2, 16/2 diagnostic checks); writable unsaved copies failed with a saved host (17/2); saved probe copies passed with and without a host (19/19, 17/17). The save state of disposable instrumented packages is the isolated fixture difference, not behavioral Search RED or a product defect. All runs left candidate XLAM hashes unchanged and Excel absent without assisted teardown; see report 026. In manual NAS review, all five November 20 add-ins were loaded and Justin found one-recording typing/filtering prompt, Backspace/clear restoring the row, and typing after selection clearing evidence and dependent actions without error. Native latency at scale remains unmeasured.

Symptom -> known fixture condition -> rule -> required proof: native Search probe cannot start because unsaved instrumented XLAMs fail Viewer caption assignment -> D18-SEARCH-INPUT-01 -> save disposable probe copies before entry, then measure native Search/Events timing on a larger library and obtain visible NAS operator confirmation.

## Do Not Repeat

Do not optimize parsing or bulk ListBox assignment from static suspicion: the reported library showed one recording, the working Viewer also uses per-row AddItem, and the hidden suite leaves no row automatically selected. Do not rerun the native Search fixture with unsaved in-memory package probes or attribute its `80010007` to product Search. Keep PR17 and PR8 draft; no acceptance claim.

## Assumptions to Re-verify

Recheck Excel startup entries, actual loaded package identity after any Excel restart, console/RDP mode, current Excel build and both PR heads. The prior slow session's actual Operations path was never captured; the current candidate file hashes were recorded before launch, not while locked open.

## Open questions and blockers

Does native Search stay responsive with a larger recording library? The test-only saved-copy correction and qualified native timing remain open. 4be-A, same-task How-To/Diagnostic/Compare, final B0, full-chain and Release 1 acceptance remain open.

## Immediate next action

Test-first, make the native Search fixture save disposable instrumented XLAM copies before Viewer entry; verify the public callback and retained checks, then run its native keyboard probe.

## Critical references

Architecture v4.11 D18-SEARCH-INPUT-01; Plan 022; Controls 1.788; report 026; guidance 027; `modInventoryViewer.OpenInventoryViewer`, `frmInventoryViewer.SetWarehouse`, `frmActionPaths.mSearch_Change`; `Slice4beViewerStartup.ps1`, `Slice4beRecordingReader.ps1`, `Slice4beRecordingSearchNative.ps1`; code PR17; docs PR8.
