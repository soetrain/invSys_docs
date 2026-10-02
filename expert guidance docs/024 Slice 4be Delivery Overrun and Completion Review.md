# Slice 4be delivery overrun and completion review

**Status:** Advisory problem report for expert review; not an architecture decision,
implementation plan replacement, approval request for deployment, or acceptance claim.
**Prepared:** 2026-10-02. **Code baseline:** main `50241047`.
**Documentation baseline:** main `1afe58c`, Architecture v4.11, Plan 022,
Controls 1.435, Production coverage audit 1.175.

## Review needed

Slice 4be has taken substantially more work than the user anticipated. A feature
described as comprehensive Event Viewer history with optional controls-based
How-To guidance has required observation infrastructure, per-control owner facts,
Action Path diagnostics, and corrections to existing workflow behavior. The recent
critical path has become Production Run correctness and verification before its
remaining observations can be integrated.

The engineering problems appear tractable. There is insufficient evidence that
the current sequence is the shortest reliable route to the approved outcome.
Independent review should produce a finite acceptance backlog, identify mandatory
correctness dependencies, and recommend how to reduce repeated verification and
documentation work without weakening Architecture v4.11 or D13.

The assistant's recent **7/10 estimate is subjective**. It has no weighted,
stable denominator and must not be interpreted as 70% of time or cost expended,
or a promise that 30% of the elapsed effort remains. The rising assertion count
does not establish feature completion. No comprehensive elapsed-time, usage-cost,
or cause-by-cause effort accounting was performed for this report.

## Approved outcome and boundaries

Architecture D18 records approval on 2026-09-07 for comprehensive Operations/Admin
coverage, dedicated Event Tracking Settings, and one versioned Action Path with
How-To, Diagnostic and Compare presentations plus personal view selection.
Authored instructions, captured observations and diagnostic results retain
distinct provenance. This supersedes the older curated-only contract.

The core boundaries are already settled:

- Events describe owner-observed facts; publication does not authorize an action.
- How-To/Diagnostic do not execute repairs, retries, overrides or business writes.
- Viewer reads published projections and the training library; opening, refreshing
  or evaluating a path does not open canonical workbooks or process inboxes.
- Core/Domain remain headless; Operations/Admin retain their respective UI owners.
- Captured workbook, warehouse and session binding, current capabilities, exact
  immutable `System_Key`, and preservation of unknown user columns remain mandatory.
- D12 packaging, D13 test-first evidence and semantic inheritance remain binding.
  A material contradiction requires an approved architecture decision, followed
  by synchronized Plan 022 and controls changes.

The prior critique in document 023 is useful context, but its sample payloads and
proposed executable Action Path classifications are not automatically requirements.
D18 explicitly declines executable Navigate/Retry/Repair/Override path types.
No new global recovery service or alternative hybrid contract is recommended here.

## Verified baseline and its limits

The current isolated candidate is `validation-production-complete-entry-02`.
It is unpromoted. The table summarizes scoped evidence, not human Release 1 acceptance.

| Gate | Verified result | What remains outside that result |
|---|---|---|
| Complete Run actual-handler baseline | 258/258; all 206 prior ordered checks retained; five instrumented XLAM compiles | Worksheet completion/interruption, exact partial-submission evidence and Complete Run observation integration |
| Check In regression on the same candidate | 629/629; prior ordered checks retained | Whole-feature acceptance cannot be inferred from this gate |
| Packaged smoke | 86/86 | Does not certify every UI workflow or NAS scenario |
| Corrected Release 1 chain | 32 chain / 48 live-role / 15 warehouse checks, all prior ordered identities retained | Does not prove every reusable Production scenario or final user acceptance |
| Layout | 18 requested size/page pairs across six pages; six captures match prior reviewed images | Empty Run List/Settings geometry does not prove populated scrolling or multiline usability |
| Latest static maintenance | 291 components, 6,177 procedures, 135,555 lines; 28 unchanged size limits; three schemas; 381 PowerShell parses | Static counts are maintenance indicators, not correctness or completion percentages |

Settings/package preservation, normal unassisted Excel closure and delayed zero
native-crash audits accompany the latest passing gates. The latest closure test
ran 03:10:09–03:24:41 UTC on 2026-10-02; independent verification completed after
the usage interruption. It was not restarted to obtain the passing result.

The four new reusable closure cases exercise the actual handler at entry,
pending UI yield, and AvailableQuantity/EntityKind return boundaries. The entry
form survives and visibly refuses the stale workbook. The three in-action cases
dismiss the form; tests do not query or recreate dismissed controls. Exact inputs,
owner state, saved workbook bytes and the unrelated workbook remain preserved.
There are no later reads or form reinitializations. This is added evidence under
an existing contract, not a new runtime correction or a manufactured RED.

Earlier failures remain part of the evidence. Later GREEN does not erase them
or establish their root cause. The latest static report also retains 9 literal
and 45 unresolved dynamic calls and 190 duplicate candidates; those figures do
not mean 45 confirmed defects or authorize automatic removal/refactoring.

## Why the work expanded

### 1. Comprehensive observations exposed real business-action defects

An honest diagnostic cannot certify an operation from a button click, a returned
Sub, status text, or a state left by an earlier attempt. It needs the current
attempt's owner result. Exercising actual handlers uncovered defects that matter
independently of tracking:

| Observed defect | Governing rule | Current disposition |
|---|---|---|
| Complete Run with no selected Process invoked a whole-run path and consumed inventory | D15 selected-Process execution | Corrected; actual-handler RED/GREEN recorded |
| Complete Run after target/session/sign-out changes still reached the owner; a replaced session consumed input | D18 captured binding | Corrected and packaged regression evidence retained |
| Binding/capability loss after pending UI or real inventory reads allowed subsequent owner/projection work | D18 continuation and Core capability boundary | Corrected; interruption tests protect the real boundaries |
| Loading/busy calls reached completion; nested calls reached the owner twice and replaced the outer success message | D18 action-entry semantics | Typed guard added; RED/GREEN, original-error propagation and separate-click recovery verified |
| Native workbook closure needed distinction between a surviving form and a dismissed form | D18 binding and actual operator opportunity | Current unchanged candidate passes the new closure tests |

These findings justify specific corrections. They do **not** imply that every
adjacent cleanup or conceivable edge case must precede the next observation.
That dependency judgment needs a finite, reviewable map.

### 2. Local success, submission and canonical application are different facts

Complete Run has reusable and worksheet-backed owners. A reusable Process may
submit consumption and output events separately. The worksheet path uses a typed
Production session and result. Failure can occur before any write, after a write
attempt without acknowledgment, after one accepted event, or during later refresh.

Current source exposes distinctions that the remaining design must respect:

- Core `QueuePayloadEventCurrent` has `writeAttemptedOut`. A generated EventId
  alone does not establish that the event was submitted.
- `CompleteReusableProcess` currently tests aggregate processor counts; those
  counts alone do not prove application of the exact submitted event.
- The worksheet path creates identities before queueing and has separate consume,
  complete, processor and refresh state. Those facts must not collapse into a
  single inferred success flag.
- Core's `InventorySourceReferences` serializes owner-supplied facts; it does not
  independently prove application. The existing worksheet action's
  `ObserveSubmission` helper currently uses Designs references and cannot simply
  be reused for Inventory events.

This is a bounded owner-evidence problem, but its complete observation contract
and partial-failure tests remain unfinished. Do not solve it by parsing messages,
assuming rollback, reading canonical authority from Viewer, or treating the
existence of an output identity as proof of successful completion.

### 3. Test infrastructure and desktop availability added a separate cost

Two recent full-chain attempts failed during canonical projection rebuild with
RPC `0x800706BE` and native Excel/ntdll `c0000028` evidence. WER reported
`OFFICE_MODULE_VERSION_MISMATCH`. This is evidence, not a verified diagnosis that
Office repair, elevation, or any particular product change is required.

The standalone ordered-live diagnostic passed both with and without tracing.
Investigation separately demonstrated that the canonical chain advanced after
Admin Quit while that Excel process was still exiting. A focused tooling test
gave 3 PASS/1 expected FAIL; an existing cleanup wait at that boundary gave 4/4.
The corrected chain then passed 32/48/15 on unchanged packages. **Stage overlap
has not been proved to cause either native crash.** Do not repeat unchanged
broad trials indefinitely or label these crashes as product behavioral RED.

Desktop Win32 error 5 is a different failure class. The user requires work to
stop and the goal to pause when actual screen-access error 5 recurs. An expired
probe, a VBA error, and a native Excel crash are not equivalent evidence. Monitors
covering the latest completion gate reported no desktop failures. There is no
claim of continuous monitoring after those timed probes ended overnight.

Tests also carry real startup, fixture, compilation and teardown costs. The latest
258-check completion run took about 14.5 minutes, and the retained 629-check
Check In run about 17 minutes. A fresh full fixture for each small addition can
consume substantial time even when runtime code does not change. No measured
aggregate fraction of the overrun is assigned to this overhead.

### 4. The delivery process needs correction too

The agent has tended to deepen the current prerequisite before returning to the
release-level completion map. That can be locally sound while producing an
unbounded sequence overall. The weak completion estimate is part of this problem.

Evidence is also repeated throughout the authority documents. At this baseline,
Architecture has 6,654 lines, Plan 022 has 10,949, and the controls catalog has
7,190. Historical paragraphs still contain superseded “pending” states alongside
newer GREEN entries. The latest records distinguish them, but the reader must
reconstruct precedence and chronology across substantial text. This increases
review and continuation cost and makes counts easy to misinterpret.

The remedy is not to delete useful evidence or weaken architectural discipline.
It is to keep normative rules and current acceptance state concise, link exact
evidence from a separate ledger, and make superseded status unmistakable. Any
reorganization must preserve approvals, traceability and unrelated user edits.

## Remaining scope that must be made finite

The Production audit records 64/68 constructed buttons wired in the catalog25
lineage, with four explicit Pending rows: Apply Scale, Complete Run, Next Batch
and Print Recall. Later entry guards and tests do not themselves add tracking
wiring. **64/68 is not 94% acceptance.** The census excludes other forms and
launchers; registered controls still have evidence and acceptance obligations.

The same audit lists 30 applicable non-button handlers, of which two Assignment
selections have observation evidence and 28 retain review work. These are not
28 automatic new event types. Some are programmatic mirrors, nested callbacks,
or text changes that must not become duplicate events or keystroke surveillance.

| Remaining workstream | Concrete missing result |
|---|---|
| Complete Run | Worksheet positive/interruption tests, exact partial-submission facts, approved/refined observation mapping, packaged tracking and independent Action Path proof |
| Apply Scale | Resolve RUN-SCALE-01; do not silently preserve a conflicting authority path |
| List/Tree Apply wording | Resolve RUN-UI-01; failed observation does not authorize a message exception |
| Next Batch / Print Recall | Define owner fact and outcome for each actual handler, then test/integrate it |
| Non-button coverage | Classify each of 28 handlers as deliberate action, excluded programmatic/helper behavior, or an explicit decision; test only the resulting contract |
| Remaining Operations/Admin coverage | Reconcile the global catalog and controls acceptance record; this report does not claim a freshly completed census outside Production |
| User-facing acceptance | Populated lists, long values, multiline/detail scrolling, both Action Path presentations and Compare, Admin settings and personal preference behavior |
| Release acceptance | Preserve all accepted role behavior and complete the applicable reusable Production, human and NAS/station evidence on the final candidate |

Two decisions are explicitly **unapproved** in Architecture:

- **RUN-SCALE-01:** With Designs enabled but no reusable run loaded, the proposal
  refuses worksheet Apply Scale before the older Design/BOM owner or staging
  mutation and directs the operator to Load Recipe. Existing released authority
  requirements do not permit an invented legacy fallback.
- **RUN-UI-01:** The proposal replaces misleading allocation-success wording when
  required worksheet staging tables/columns are unavailable. Four assertions
  express this proposal; they are not approved acceptance gates yet.

Other user approvals must not be treated as blanket approval of these decisions.
An expert may recommend a resolution; the required architectural approval remains
separate before conflicting behavior is implemented.

## Recommended completion approach for review

These are recommendations, not enacted changes to Plan 022 or required gates.

1. **Create one current acceptance matrix.** Each remaining item names the rule,
   owner, actual callback, approved outcome, smallest protecting test, dependencies,
   candidate/package hashes, visual proof and acceptance state. Registration,
   automated verification and human acceptance must be separate columns. Replace
   subjective percentages with closed/open milestone counts and explicit unknowns.
2. **Bound correctness prerequisites.** For each newly discovered issue, identify
   the exact approved release requirement it blocks. Record nonblocking cleanup
   separately. Escalate contradictions immediately rather than letting them become
   another long local test sequence. Do not defer an approved requirement by fiat.
3. **Resolve the two named contract decisions.** Obtain concrete decisions on
   RUN-SCALE-01/RUN-UI-01 while independent approved work continues. Avoid revisiting
   already approved connection-write, UOM, Auth or Event Detail decisions merely
   because the session changed.
4. **Finish Complete Run as a bounded deliverable.** Agree the owner-fact model,
   enumerate submission/partial-failure outcomes and needed worksheet cases, then
   implement observations and prove one independent recorded/curated diagnostic
   path. Preserve the current reusable baseline instead of repeatedly rebuilding it.
5. **Classify remaining controls in a batch before coding.** Review Next Batch,
   Print Recall and the non-button census together so the true remaining size is
   visible. Implementation can still use small D13 slices with separate evidence.
6. **Optimize test selection, not test obligations.** Keep real packaged callbacks,
   meaningful RED/GREEN and every required release gate. Consider focused sub-gates
   for new test-only cases while preserving the aggregate suite. Tie reruns to
   changed runtime/package hashes, failures or an identified dependency; do not
   rerun unchanged accepted smoke/layout/chain merely because evidence prose changed.
   If a proposed cadence conflicts with D13, resolve it explicitly before adopting it.
7. **Use explicit escalation and finish rules.** After a repeated unexplained
   native failure, preserve receipts and run a bounded discriminating experiment;
   do not cycle broad retries. Set a review point by completed milestone or elapsed
   work, then reassess scope and blockers. The exact limits should be agreed rather
   than invented as new project policy in this report.
8. **Finish with visible acceptance on the final candidate.** Reserve an unlocked
   desktop window for user comparison and NAS/station work. Give each open visual
   requirement a concrete pass/fail task. Do not infer usability from a screenshot
   hash or final release acceptance from a test count.

## Questions for the expert

1. Which remaining correctness checks are mandatory prerequisites to trustworthy
   Complete Run observations, and which proposed checks can be consolidated while
   preserving the approved invariants and D13?
2. What is the smallest explicit owner-fact interface that represents partial
   Inventory submissions, uncertain acknowledgment and exact application without
   turning the tracking layer into business authority?
3. Does the proposed remaining-control classification close D18's comprehensive
   coverage obligation without recording helpers, mirrors or keystrokes?
4. What resolutions should be proposed for RUN-SCALE-01 and RUN-UI-01? Name the
   exact architectural text that must change if recommending different behavior.
5. Can required evidence be scheduled by dependency and frozen candidate so small
   test-only expansions avoid unnecessary broad reruns? Specify obligations retained.
6. Does the native-crash evidence warrant a separate bounded environment diagnosis
   now, or is retaining it with passing current gates sufficient pending recurrence?
   Distinguish a hypothesis from a confirmed cause.
7. What finite milestones would make a remaining-effort estimate defensible, and
   what evidence is still missing to estimate them?

The requested review output is a prioritized dependency table, a list of exact
decisions requiring approval, and acceptance criteria for the last deliverables.
A general recommendation to “add more testing,” an unbounded rewrite, or automatic
scope reduction would not resolve the delivery problem.

## References and evidence access

Primary authority and current acceptance records:

- [Architecture v4.11](../0%20plan%20docs/xlam_invSys/invSys-Design-v4.11.md): D13,
  D15, D18; named RUN-SCALE-01/RUN-UI-01 proposals.
- [Plan 022](022%20Deployed%20Operations%20Launcher%20and%20NAS%20Runtime%20Stabilization%20Plan.md).
- [Controls catalog](../0%20plan%20docs/xlam_invSys/invSys-Controls-v1.md).
- [Production coverage audit](../0%20plan%20docs/xlam_invSys/invSys-Production-Tracking-Coverage-v1.md):
  constructed-button and non-button tables; distinguish historical status paragraphs.
- [Prior Slice 4be critique](023%20Slice%204be%20Critique.md): user-owned local file,
  currently untracked; preserved without editing or adding it to this commit.

Code repository at `50241047`:

- `tests/integration/plan022_slice4be_production_complete_results.md`: RED/GREEN,
  failed native attempts, corrected chain and current closure evidence.
- `tests/tooling/Slice4beProductionCompleteBaseline.ps1`,
  `Slice4beProductionCompleteClosed.ps1`, `Test-Slice4beChainStageExit.ps1`;
  `tools/validate_release1_full_chain.ps1`.
- `src/Production/Forms/frmProduction.frm`: `mBtnManagerApplyOutput_Click`,
  `CompleteProductionRun`.
- `src/Production/Modules/modProductionCompleteActions.bas`,
  `modProductionReusableRun.bas` (`CompleteReusableProcess`),
  `modProductionCompletionService.bas` (`QueueProductionSessionEvents`,
  `ExecuteProductionSession`), and `mProduction.bas`
  (`CompleteProductionRunAfterCheckInForOutput`).
- `src/Production/ClassModules/cProductionWorksheetAction.cls`,
  `cProductionLifecycleFacts.cls`, `cProductionCompletionResult.cls`;
  `src/Core/Modules/modRoleEventWriter.bas` and `modActivity.bas`.

Local ignored receipts, not committed runtime data:

- `reports/runtime/production-run-local-controller/9f7cf60d4d554fa8a4239d77737b42fe/verification.json`
  and `native-closure-verification.json`.
- `reports/runtime/complete-entry02-regression/chain-a98044117e04412d98d6d30a403bc752/verification.json`.
- `reports/runtime/complete-closed-static-01/ratchet-verification.json`.

This report contains sanitized findings and identifiers only. An external reviewer
will need the referenced source and sanitized evidence excerpts; the local runtime
folders and untracked critique will not appear automatically in a GitHub checkout.
No runtime reports, operational workbooks, screenshots or credentials are attached.
