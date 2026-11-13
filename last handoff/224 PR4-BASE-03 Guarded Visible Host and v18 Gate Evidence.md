# Handoff 224: PR4-BASE-03 guarded visible host and v18 gate evidence

Date: 2026-11-11 Pacific (series closed 2026-11-11T03:58Z). Predecessor: handoff 223.

## Goal and release outcome

Complete invSys Release 1 Slice 4be-A under Architecture v4.11 and Plan 022. v18 qualified five Inventory recording/guide selectors and focused Tester/Create recording; v17.1 through v17.4 isolated the visible-host window fault. Run 07 Recording.Uom remains unqualified; none of it establishes broad Inventory GREEN, 4be-A or release acceptance, and the Excel native-exit defect remains unresolved.

## Current verified state

Last verified: 2026-11-11 Pacific. Code branch work/d8-inv-edit-receipts is at merge commit 550d9ff08fcce2d78c1a059a51ca7065928c9c66 (draft PR4 against main). Code PR10 merged 55f46fcc60fe90c3371be91e3a33cb7078bc0512 from work/d8-guide-host-visible into that working branch; it changes only tests/tooling/Slice4beInventoryGuide.ps1, Slice4beAdminTesterRecording.ps1 and Slice4beAdminCreateRecording.ps1. Docs branch work/perplexity-handoff218 was at 2673341f before this handoff (draft PR5); Plan 022 and Controls 1.734 carry the matching evidence-only checkpoint. Worker worktree invSys_edit_receipts is back on work/d8-inv-edit-receipts at 550d9ff0 with no tracked changes. The merge commit tree is identical to the v18-tested 55f46fcc tree; the approved PR4-BASE-03 manifest still names worker commit 55f46fcc, so future runs on that baseline must detach there. No VBA, XLAM, architecture or product contract changed. The main code checkout still has Justin's unrelated modified AGENTS.md; the docs checkout keeps its pre-existing unrelated changes listed in handoff 223. Preserve them. At the v18 delayed audit Excel was closed, settings restored and the 538 PR4-BASE-03 inputs preserved; recheck before another test or build.

## Decisions and constraints

- On 2026-11-11 Justin authorized PR10 integration; it was marked ready and merged into work/d8-inv-edit-receipts as 550d9ff0. PR4 remains draft against main. Docs PR5 remains draft and unmerged.

- Justin approved a PR4-BASE-03 re-pin exception (2026-11-10) for exactly the three tooling files above, solely to carry the visible-host fix, as a separate manifest with PR4-BASE-02 preserved unchanged. PR4-BASE-03 manifest SHA256 B03E9264DB8789A233A55BE312B6F1BAC58D49464EB182E07E3644B2199A53E3 (538 pins, worker commit 55f46fcc) was approved for runtime qualification the same day. PR4-BASE-02 F2DE6B82..., PR4-BASE-01 6FDE11C2..., the worker file AD8797E8... and all five XLAM pins are unchanged.
- The re-pinned hashes are the on-disk values after checkout (the worker repo has core.autocrlf=true, so git writes these files with CRLF): Guide 196B2F47..., TesterRecording B8BE171A..., CreateRecording F150E6C3.... An earlier BASE-03 draft that pinned LF blob hashes (8E38D215...) is kept as *.superseded-lf-blob-hashes.json; it was never approved or used. Returning the worktree to the branch rewrites the Guide with CRLF, so its on-disk hash (1EC53A9E...) no longer matches the PR4-BASE-02 pin EB81A925... of the old hand-written LF copy; recorded PR4-BASE-02 evidence is unaffected.
- The change is test tooling, not the invSys product contract, so no Architecture amendment was needed. v17.1 through v17.4 ran on tooling copies outside the pinned tree; only 55f46fcc changed pinned files, under the exception.
- v18 introduced the rule that a failed check-identity or speed gate is recorded and the series continues to the next run; evidence, restoration, preservation, Excel-ownership, desktop and policy faults still stop the series. One owner approval covered the seven-run series; worker detach to 55f46fcc for the series was named in it.
- v18 stands alone. It is not added to v17, v16's 772 or v13.1's passes, and is not Inventory GREEN, 4be-A or Release 1 acceptance. Codex executed; Perplexity built, reviewed and audited.

## Evidence and traceability

- v17 (PR4-BASE-02, 2026-11-10, reports/runtime/inventory-gate-v17-pr4-base-02-20261110/): runs 01 WorksheetResults 64/64 (62.8 min, attended after three native tracking notices were acknowledged), 02 OwnResults 50/50, 03 HeadlessWorksheet 56/56, 04 Recording 151/151, 05 Recording.Selections 65/65 exact with natural exits; run 06 Recording.Selections.Guide stopped the series at 9 PASS, 1 FAIL, 66 missing (worker exit 1). Runs 07 Recording.Uom (68), 08 and 09 did not run in v17.
- v17.1 to v17.3 (run-06 only, tooling copies; each 9 PASS, 1 FAIL): the pinned guide's line 39 (`Windows.Item(1).Visible=$true; Activate(); $excel.Visible=$true`), run against an already visible and active host, moved Application.Hwnd to a new frame and left Workbook.Windows and Application.Windows empty for 20 s; line 40 then failed with DISP_E_BADINDEX. Showing Excel before Workbooks.Add gave a healthy host.
- v17.4 (reports/runtime/inventory-gate-v17-4-guide-host-20261110/): run-06a isolated the statement; HostWindowsCount dropped to 0 only after statement 3, `Application.Visible = True` re-asserted on visible Excel. Run-06b, with each statement executed only when its target was not already in the requested state, passed 75/75 exact with a natural exit.
- 55f46fcc carries that guard into the three pinned guide/recording host files (PR10). PR4-BASE-03 pins the result.
- v18 (PR4-BASE-03, reports/runtime/inventory-gate-v18-pr4-base-03-20261110/, launcher EF020A5F..., expected checks C93A6573...): 04 Recording 151/151, 05 Recording.Selections 65/65, 06 Recording.Selections.Guide 75/75, 08 Recording.Uom.Guide 74/74, 09 Recording.Guide 73/73, T1 focused Tester.Recording 163/163, C1 focused Create.Recording 170/170; 771 identities exact, seven workers exited 0, seven natural Excel exits (stdout-quiet to exit 1.4 to 6.9 s; these selectors write no AfterQuit trace), settings restored, inputs preserved, no termination, delayed audit zero Excel fault events. Series ran 03:08Z to 03:58Z of a 240-minute budget. Perplexity re-compared each worker green.json against the expected file by name (audit-v18.md in the series folder). T1/C1 expected names came from the 2026-10-14 all-PASS Tester (c6ea2401...) and Create (3777d32b...) reports restricted to the focused selector and reproduced by a static model of the 55f46fcc tooling.
- Correction to the v18 handoff text: it said run 07 passed in v17. It did not run; Recording.Uom remains unqualified on both baselines.

## Do Not Repeat

Do not rerun v17.1 to v17.4 to re-establish the window fault; it is isolated and the guarded workaround is proven on this host. Do not modify frozen manifests, replace the superseded BASE-03 draft, or normalize line endings in pinned files. Do not combine v18 with v17, v16 or v13.1, claim the native-exit defect fixed, or infer that the product itself re-asserts Application.Visible; that is unexamined. Justin's Microsoft report is submitted; no follow-up dialogue is open. Do not merge or promote PR4 to main; PR10 has already merged into the PR4 working branch.

## Assumptions to Re-verify

Before runtime work, verify Excel closure, desktop input, package/manifest pins, current branch and PR heads, the worker worktree position (branch and clean), settings restoration and no unfinished worker. Any new series on PR4-BASE-03 must begin with the worktree detached at 55f46fcc and must re-verify the three on-disk hashes against the manifest. The 62.8-minute v17 run 01 and the Execute slowdown remain unexplained.

## Open questions and blockers

Run 07 Recording.Uom (68 identities) has never run on PR4-BASE-02 or PR4-BASE-03. Whether the product VBA re-asserts Application.Visible on a visible Excel anywhere is unexamined. Complete retained-formatting caller coverage, broad Inventory GREEN, early Receiving B0, live-role, restart, full Release 1 chain, final compile/layout/static and visible-operator evidence, and 4be-A acceptance remain open. The Excel native-exit defect remains unresolved. PR10 is merged. Docs PR5 remains draft; CURRENT points to 224 locally but has not been published.

## Immediate next action

Run 07 Recording.Uom as a one-run series on PR4-BASE-03 under its own approval, then decide whether a bounded read-only audit of the product for Application.Visible re-assertion is warranted.

## Critical references

Architecture v4.11 D13; Plan 022 (2026-11-10 PR4-BASE-03 checkpoint); Controls 1.734; tests/tooling/Slice4beInventoryGuide.ps1, Slice4beAdminTesterRecording.ps1, Slice4beAdminCreateRecording.ps1 at 55f46fcc; PR4-BASE-01/02/03 manifests under reports/runtime/pr4-base-0x-*/; reports/runtime/inventory-gate-v17-pr4-base-02-20261110/, inventory-gate-v17-1..4-guide-host-20261110/ and inventory-gate-v18-pr4-base-03-20261110/ (incl. audit-v18.md); draft code PR4, merged PR10 and draft docs PR5; Microsoft SDI guidance https://learn.microsoft.com/en-us/office/vba/excel/Concepts/programming-for-the-single-document-interface-in-excel.
