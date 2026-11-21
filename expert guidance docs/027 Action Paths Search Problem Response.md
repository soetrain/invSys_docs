## Codex job: fix Action Paths Search by reusing the working Viewer’s in-memory search pattern

### Goal

Make **Action Paths → Search recordings** respond immediately while typing, comparable to the existing **Viewer → Events → Search** in the same Excel session.

Justin recalls arrays being used successfully to resolve earlier search slowness. The working Viewer source supports that implementation pattern: it parses loaded data into a two-dimensional Variant array retained in `mRows`, then invokes `RenderRows` when Search changes. Reuse the relevant working pattern rather than redesigning Search. The exact historical fix commit has not been identified.

### Approved scope

`D18-SEARCH-INPUT-01` is already approved. No additional architectural proposal is needed for investigation or a correction within that contract.

Preserve these requirements:

- Open/Refresh obtains recording headers through the authorized Core boundary.
- Each search edit immediately filters the previously loaded headers locally.
- No per-character Config read, NAS journal read, publication, or keystroke recording.
- Each edit checks the captured session/warehouse locally.
- Filtering clears the selected recording’s evidence, evaluation/binding, and dependent-action access.
- Explicit recording selection and other Core-backed actions retain current-policy and authority checks.
- Refresh reloads current headers under current policy, then reapplies the existing search text.
- Session/warehouse changes invalidate the old snapshot.
- Preserve exact recording identity, existing case-insensitive matching semantics, source bytes, ControlIds, and displayed wording.

Do not substitute Enter-to-search, a new Search button, or delayed filtering for the approved immediate-search behavior.

### Starting point and references

Last reviewed code commit: `aa44b245`, draft code PR17. Recheck the current head before beginning.

- [Code PR17](https://github.com/soetrain/invSys_fork/pull/17)
- [Docs PR8](https://github.com/soetrain/invSys_docs/pull/8)
- [Working Viewer source](https://github.com/soetrain/invSys_fork/blob/aa44b245/src/Operations/Forms/frmInventoryViewer.frm)
- [Current Action Paths source](https://github.com/soetrain/invSys_fork/blob/aa44b245/src/Operations/Forms/frmActionPaths.frm)
- [Performance report 026](https://github.com/soetrain/invSys_docs/blob/work/d18-picker-save-repro-docs/expert%20guidance%20docs/026%20Action%20Paths%20Search%20Native%20Performance%20Problem%20Report.md)
- Handoff 255: `last handoff/255 Native Action Paths Search Usability RED.md`
- Architecture v4.11: `D18-SEARCH-INPUT-01`
- Existing reader fixture: `tests/tooling/Slice4beRecordingReader.ps1`

## 1. Understand the remaining implementation difference

The working Viewer:

```text
Authorized data load
  → parse payload into two-dimensional dataRows array
  → retain model in mRows

Search Change
  → RenderRows(searchText)
```

The reviewed Action Paths candidate:

```text
Authorized Open/Refresh
  → retain serialized header string in mLoadedRows

Search Change
  → ContextValid
  → ClearSelectedRecording
  → split cached string into rows again
  → search each row
  → split matching rows into fields again
  → ListBox.Clear
  → AddItem and individual column assignments
```

The current Action Paths handler no longer directly calls `ListPaths` per character. Do not repeat the old per-character NAS-read diagnosis as the established cause of the new candidate’s native slowness.

Read the complete Viewer `RenderRows` path and its relevant helpers—not merely the one-line Change handler. Determine which parts can be reused while preserving Action Paths’ distinct selection and authorization boundaries.

## 2. Verify the exact build under test

Before attributing Justin’s observation to a particular implementation:

- Verify the exact loaded paths and identities of all five XLAMs.
- Establish whether the loaded Operations package is the intended validation candidate.
- Check for an older registered or loaded copy supplying the form.
- Record Excel build, loaded recording-header count, and console/RDP display mode.
- Preserve unrelated worktree changes and user workbooks.
- Follow existing Excel-ownership, desktop-input, package-build, and natural-cleanup rules.

Report 026 records different Operations hashes for the validation candidate and `deploy/current`; the exact package loaded in the slow review session was not verified. Resolve this ambiguity without assuming Justin used the wrong build.

Do not overwrite `deploy/current`, the NAS deployment, or the retained review candidate.

## 3. Establish a focused native performance baseline

Follow D13 test-first. Existing programmatic `.Value` tests are useful functional coverage but do not establish native typing responsiveness.

Measure comparable searches in the same Excel/session/display state:

1. Action Paths Search.
2. Viewer Events Search.
3. The existing responsive picker, where practical.

For Action Paths, distinguish:

- First character after selecting a recording.
- Subsequent characters with no recording selected.
- Backspace and clearing Search.
- A query retaining many matches versus one returning few or no matches.

Instrument these boundaries separately:

```text
Search Change entry/exit
ContextValid
ClearSelectedRecording
Matching/parsing
ListBox replacement/population
Selection-handler reentry
Activation/layout callbacks occurring during the interaction
```

Measure native key-to-visible-result latency separately from handler execution time. Record counts, durations, and event identities only; do not log search text or operational row values.

Identify the first measured expensive boundary. Do not blame NAS, RDP, MSForms repaint, stale packages, or string parsing without evidence.

## 4. Prioritize an array-based rendering comparison

The first implementation comparison should replace repeated parsing and per-row control updates with the existing project’s retained-array pattern:

### At authorized Open/Refresh

- Parse the complete loaded header payload once.
- Retain a two-dimensional in-memory row array containing exact recording IDs and displayed fields.
- Optionally retain a search string per row if useful, preserving the existing search predicate.
- Keep this model bound to the captured session/warehouse.
- Replace or invalidate it appropriately on Refresh failure, context change, and form lifecycle transitions.

### On each search edit

- Perform the required local context check.
- Invalidate the previous selection/evidence and dependent-action access.
- Filter from the complete retained array—not the currently displayed subset.
- Build a correctly sized two-dimensional result array.
- Test one bulk ListBox assignment:

```vb
mPaths.List = filteredRows
```

- Handle zero matches explicitly.
- Suppress selection events during list replacement and reliably restore the suppression flag on errors.
- Leave no automatically selected recording after typing.
- Keep the search textbox’s focus, caret, and entered text intact.

Avoid repeated `Split` operations over the entire serialized payload and repeated `AddItem`/per-cell updates in the hot path if the array comparison demonstrates improvement.

This is a candidate correction, not a predetermined root-cause conclusion. If measurements instead locate the delay in selection cleanup, activation, or another callback, correct that boundary rather than optimizing string matching to hide unrelated work.

## 5. Protect selection and child-form behavior

Do not clear only the visible evidence textbox.

Audit and protect:

- `mSelectedId`
- `mBinding`
- Saved evaluation identity and displayed evaluation
- Core selected-recording state
- Expectation/guide editor relationships
- Published-guide pairing and action-path child views
- Dependent button enablement

The first edit may require invalidating dependent state. Subsequent edits must not repeat unnecessary cleanup after that state is already invalidated. Any optimization must preserve the approved evidence-clearing and stale-authority behavior.

## 6. Required tests and acceptance criteria

Extend the protecting tests before product edits.

### Functional and authority checks

- Native typing filters immediately and correctly.
- Paste, Backspace, deletion, no matches, and clearing Search work.
- Backspace/clear restores rows from the complete loaded snapshot without a NAS reload.
- Exact IDs remain associated with their displayed rows.
- Existing case-insensitive matching behavior is retained.
- No per-key Config or NAS journal read occurs.
- Typing clears previous evidence/bindings/actions even when the previously selected row still matches.
- List replacement does not trigger an unintended evidence read or automatic reselection.
- Explicit selection rechecks current policy and authority.
- Sign-out, session replacement, and warehouse changes refuse stale reuse.
- Refresh loads current headers and reapplies Search.
- New recordings do not appear in the old snapshot until authorized Refresh.
- Policy/store failures cannot turn cached headers into authorization.
- Source bytes and unrelated user data remain unchanged.

### Performance and runtime checks

- Define a concrete native responsiveness threshold before evaluating GREEN; do not select it afterward to fit the result.
- Compare the baseline and correction using the same dataset, loaded build provenance, and session conditions.
- Show first-character and subsequent-character results separately.
- Require visible NAS operator retest, not only programmatic textbox assignment.
- Retain applicable packaged compile, reader, layout, integrity, and regression gates.
- Verify natural Excel closure, restored test settings, and preserved package/source hashes.

The prior `130/130` packaged result remains functional evidence. It does not override Justin’s native usability RED or establish acceptance of this correction.

## 7. Return a concise engineering handoff

Report:

1. Exact source commit and loaded package identities.
2. The measured slow boundary.
3. How the working Viewer implementation differs.
4. Baseline versus corrected native timing, including row counts and first/subsequent characters.
5. Exact files changed and why.
6. Test-first RED/GREEN results and retained regression results.
7. Whether Justin’s visible NAS retest passed.
8. Any remaining unknowns.

Keep the PR draft until the applicable gates are satisfied. Do not merge, promote packages, or claim 4be-A/Release 1 acceptance from this focused fix.

The objective is straightforward: reuse the project’s working in-memory search/rendering pattern, remove measured unnecessary work from the keypress path, and prove the actual typing experience is fast.