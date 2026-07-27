v4.11 already codifies D12 (Operations packaging) and D13 (test-first) exactly as discussed, which changes the slice plan in two ways: the plan must now be explicitly scoped to Receiving+Production+Shipping together (not just Production), and every slice touching Core/Domain/service logic must build in a RED step per D13 rather than treating tests as a later verification pass.

## Revised Slice Plan: Operations Stabilization (Receiving + Production + Shipping)

**Slice 0 — Scanner, manifest, and D13 baseline (prerequisite).** Build the source/contract scanner across Receiving, Production, and Shipping code (not Production alone), generating the implementation manifest (module sizes, `Application.Run` sites, table/column contracts, Ribbon `onAction` targets, test coverage per public entry point). Simultaneously, since D13 requires a named test protecting every contract touched, use this scanner to identify which existing Receiving/Shipping/Production functions currently have zero test coverage — that list becomes the RED backlog for Slices 1–4, not an afterthought.

**Slice 1 — Define role-agnostic RunSession pattern, starting with Production.** Build the typed `ProductionRunSession` per the original design (recipe loaded → allocated → checked in → completion submitted → processor applied → refreshed → ready for next batch), keyed by (ITEM_CODE, Location) per D9/D10's existing SKU+Location projection contract. Per D13, write the failing session-state test before implementing the session object. Because Receiving and Shipping need the equivalent pattern in Slice 8, design the session object's shape (progression states, structured result, not string reports) so it generalizes — e.g., `RunSession` base behavior with role-specific state enums — rather than building something Production-only that Receiving/Shipping later have to reinvent.

**Slice 2 — Disable implicit legacy recipe fallback under Designs (Production-specific).** Unchanged from before: DesignsEnabled=True means Production reads released Designs Domain recipes only, per the now-spec'd `tblDesigns`/`tblDesignLines` projection contract in v4.11's new Designs Domain Tables section. RED test: assert legacy fallback does not fire when Designs is enabled, written before the fallback is removed.

**Slice 3 — Extract completion into one typed service (Production), with the RED test as the gate.** Per D13's explicit completion prohibition, this slice is not allowed to be marked done if the service was extracted first and tests written to confirm it worked — the packaged two-batch form-action test target must be written and observed failing against the *current* code before the extraction begins.

**Slice 4 — Make the Production form a thin renderer/controller.** Unchanged; depends on Slice 3's service existing.

**Slice 5 — Replace same-XLAM Application.Run with direct calls (Production).** Now scoped inside the anticipated `invSys.Operations.xlam` project boundary — this is a good moment to also decide whether Receiving/Shipping have the same same-project `Application.Run` anti-pattern, since D12 explicitly warns that packaging alone must not hide latent callback ambiguity.

**Slice 6 — Two-batch Production form-action integration test.** This is now explicitly the D13-mandated RED artifact for Slice 3/4, not a separate later step — v4.11 folds this into Phase 6 tasks directly ("Write and record RED for a packaged two-batch Production form-action test before refactoring the Production UI/run-session wiring"). Sequence it *before* Slices 3–4 complete, not after.

**Slice 7 — Anchor-based layout replacement (Production).** Unchanged, parallelizable.

**Slice 8 — Prove and migrate Receiving and Shipping onto the same RunSession/service pattern.** This is the slice your last correction actually demanded and the original plan under-weighted. Concretely:
- Apply the Slice 1 RunSession pattern to Receiving (received-qty staging → Confirm Writes → processor applied → snapshot refreshed) and Shipping (the A+B event loop already spec'd in D11 — Add/Remove/To Shipments/Shipments Sent staging → processor applied → NAS Inv refreshed).
- Extract Receiving's and Shipping's completion/posting logic into typed services mirroring Slice 3's Production service, since D13 will require RED tests for these before implementation regardless.
- Write packaged form-action integration tests for Receiving (Confirm Writes path) and Shipping (Shipments Sent path) analogous to Slice 6, per D13's requirement that high-risk form-action paths get integration-level RED tests, not unit tests alone.
- Retire whichever legacy Receiving/Shipping worksheet-mutation paths the domain-engine loop makes redundant — this was your explicit requirement from earlier in the conversation and needed its own slice rather than being folded silently into packaging.

**Slice 9 — Operations packaging per D12.** Sequenced after Slice 8, not before, since D12's failure-isolation rule requires that Receiving, Production, and Shipping modules each compile and initialize cleanly before being combined into one XLAM — consolidating a still-unstable Production module into the shared binary risks D12's stated regression risk (one role's failure taking down all three). This slice executes the now-fully-spec'd D12 requirements directly: legacy XLAM retirement/coexistence prevention, capability-gated ribbon groups, package-manifest validation (exactly five XLAMs), duplicate-callback/tab tests, and selective complete-project builds.

**Slice 10 — Full Receiving → Production → Boxing → Shipping chain validation.** The closing integration slice validating the end-to-end flow against the consolidated Operations package, which is the actual completion criterion from your original assessment ("efficient route to completing Receiving → Production → Boxing → Shipping reliably").

## What changed from the original plan

| Aspect | Original slice plan | Revised plan |
|---|---|---|
| Scope | Production-centric; Receiving/Shipping treated as a late addendum (old Slice 9) | Receiving and Shipping get their own full RunSession/service/test treatment (Slice 8), sequenced as core work, not an afterthought |
| Test ordering | Slice 6 (integration test) came after Slices 3–4 (extraction) | D13 requires the RED test before extraction; Slice 6's test now gates Slices 3–4 rather than following them |
| Packaging timing | Packaging (old Slice 8) preceded full Receiving/Shipping proof | Packaging (Slice 9) now follows Slice 8, since D12's failure-isolation rule requires each role module proven stable before consolidation |
| Legacy retirement | Implied but not a distinct step | Explicit sub-task inside Slice 8, matching your stated requirement and D12's coexistence-prevention rules |

This ordering also satisfies the "don't let AI drift" concern directly: v4.11 now contains a binding rule (D13's session-start check) requiring any session touching this work to name the contract and its protecting test before editing code, which functions as the standing-goal anchor discussed earlier — the slice plan no longer depends on a chat session remembering to ask "did we write the test first?"