# Handoff 244: standalone Training onboarding proposal cancelled

Date: 2026-11-15 Pacific. Predecessor: handoff 243.

## Goal and release outcome

Enable the new Training warehouse hands-on 4be-A review under Architecture v4.11 D8/D18; no warehouse or human acceptance is complete. The proposed first-use contract was cancelled by Justin.

## Current verified state

Last verified 2026-11-15: Justin cancelled D8-TRAINING-ONBOARD-01. Architecture v4.11, Plan 022 and Controls 1.758 were restored to their pre-proposal content; no new contract is approved or implemented. Code PR16 branch `work/d14-packaged-chain-coverage` at `6e866f16210ac864a891a3b425e2d99001b19d1a` contains the operator review sheet and is tracked-clean. Main code checkout has Justin's unrelated `AGENTS.md` edit; docs checkout also has unrelated pre-existing edits/untracked files. Excel is closed. The ignored review candidate's five XLAM hashes matched PR4-BASE-18, and `TrainingHub` remains empty.

## Decisions and constraints

Justin chose a new disposable Training warehouse and approved a **test-only** post-create PIN and `ACTION_PATH_MAINT` grant for preview. He then cancelled the proposed user-side standalone local first-use contract. Existing D8-CREATE-01 requires captured signed-in source Admin and remains binding. For NAS deployment, authenticated Windows SMB write access is a separate layer; it does not authorize the invSys Create form. Do not infer test-only grant as product acceptance.

## Evidence and traceability

Source: `modAdminCreateEntry.OpenFromRibbon` captures the current source context; `modWarehouseCreateCommands.CreateWarehouse` rechecks source ADMIN_MAINT. `modWarehouseBootstrap` creates the initial user with ADMIN_MAINT/RECEIVE_POST but no PIN hash; normal sign-in rejects blank hash. Users & Roles exposes six rights and cannot grant `ACTION_PATH_MAINT`. Symptom: a new user cannot create a warehouse through the approved interactive path or author a guide after current bootstrap. Governing rules: D8-CREATE-01, D8-USERS-01, D18-REPLAY-01. No behavioral RED/GREEN has run for the proposed correction.

## Do Not Repeat

Do not implement the cancelled standalone first-use proposal, remove the D8 source check from the existing Create command, treat Windows/NAS identity as an invSys session, use an unsecured warehouse for operations, or count a fixture-only preview as onboarding GREEN. Do not edit or publish operational workbooks to simulate the path.

## Assumptions to Re-verify

Recheck current PR heads, Excel closure, candidate hashes and empty Training hub before any runtime setup. The cancelled first-use setup claim and pending-state design were not implemented. NAS write access has not been tested in this session.

## Open questions and blockers

No first-use user-side path currently exists. The signed-in Create path still needs a captured source Admin session; the initial new Admin has a blank PIN and no ordinary guide-right grant. Justin has not named a source Admin session. NAS deployment can use existing Windows SMB access if that session has write permission; otherwise the owner must establish share credentials privately in Windows. This does not solve the Create authority gap.

## Immediate next action

Answer the NAS credential question, then keep 4be-A preview setup within the current D8-CREATE-01 contract or obtain a new explicit architectural direction before product changes.

## Critical references

`0 plan docs/xlam_invSys/invSys-Design-v4.11.md` D8-CREATE-01 and D-NAS; Plan 022; Controls 1.758; code PR16; docs PR5; `tests/integration/plan022_slice4be_operator_review.md`; ignored `reports/runtime/4be-a-human-review-20261115/` candidate and cancelled local proposal.
