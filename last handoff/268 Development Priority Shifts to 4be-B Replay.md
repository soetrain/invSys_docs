# 268 Development priority shifts to 4be-B replay

## Goal and release outcome

Justin directed on November 25, 2026 that development move to 4be-B, Replay and proof, starting in the next branched chat. The first B1 case is replay of Admin Seed Demo Inventory in a newly generated Training warehouse.

## Current verified state

Last verified November 25. The code review branch `work/d18-picker-save-repro` is at `ca190e4a`. The prior docs checkpoint is `3bef92ea` on `work/perplexity-handoff218`, published to `work/d18-picker-save-repro-docs`. This transition updates project records only. B1 implementation and its first RED are pending. Existing unrelated local edits remain in place.

## Decisions and constraints

Justin reported that search worked as expected and the controls he tried worked. The development priority changes now. Formal 4be-A, 4be-B and Release 1 acceptance remain open. Retain the existing Receiving B0 replay as a regression.

## Evidence and traceability

Handoff 267 identifies the five tested package hashes and the 121/121 Receiving B0 result. The B1 test must link the seed outcome to the current replay through the ordinary packaged Admin handler. Matching pre-existing inventory alone is insufficient.

## Do Not Repeat

Use the existing Receiving replay proof and shared contracts. Avoid restarting B0 or treating setup, package-loading or compile errors as behavioral RED.

## Assumptions to Re-verify

Check current source, package identities and Excel ownership before a new packaged test. The build and individual controls Justin exercised during his visible review were unspecified.

## Open questions and blockers

The complete 4be-A onboarding walkthrough, picker no-match explanation and remaining recording coverage stay open under their existing records.

## Immediate next action

In the next branched chat, run a focused B1 test that fails when Admin Seed replay cannot prove this run produced the inventory outcome in the new Training warehouse.

## Critical references

Architecture v4.11, 4be-B and D18-REPLAY-01. Plan 022, B1 Replay and Admin. Controls v1, November 25 operator feedback. Handoff 267 and `tests/tooling/Test-Slice4beReceivingReplay.ps1`.
