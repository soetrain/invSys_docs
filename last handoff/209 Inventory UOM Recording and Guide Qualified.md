# 209 Inventory UOM Recording and Guide Qualified

## 1. Goal and release outcome

Complete4be-A Recording/Event Viewer/How-To and required early Receiving B0 under
Architecture4.11/Plan022. **4be-A remains incomplete**; broader4be-B/fullR1 are
separate. Inventory UOM prompt recording and independent guide evidence now qualify.

## 2. Current verified state

Last verified2026-10-21 07:08Z. Main code remains main/origin/main `6d247543`.
Implementation is sibling `../invSys_edit_receipts`, branch
`work/d8-inv-edit-receipts`, pushed `bd438789`, draft PR4 targeting soetrain/main.
Docs main is pushed `cdfb8ad` before this handoff. Architecture/Plan022/Controls
1.713 agree; catalog46 has199 controls. Candidate main
`deploy/validation-inventory-uom-02` is unpromoted.

Draft code is clean. Preserve unrelated main AGENTS.md and docs changes: deleted
old024, modified handoff067, untracked guidance023/renamed024/025. Excel, test
controllers and desktop monitor are closed; normal cleanup/audits are confirmed.

## 3. Decisions and constraints

No proposal pending. D18 semantic inheritance permits this observation refinement.
`ADMIN_INVENTORY_UOM_ADD` reports STAGED for a new/existing local choice, CANCELLED
for empty/cancel, REFUSED for the existing post-dialog owner refusal. These assert
no canonical change; only STAGED is positive CommandCompleted. Lost recording
context may leave an incomplete attempt. No entered value or canonical reference;
automatic combo updates are excluded. All198 prior definitions are unchanged.
The guide editor now exposes its catalog-defined REFUSED outcome.

Keep packaged callbacks/handlers, Admin-generated fixtures, captured context and
current authority intact. No rebuild with Excel open; wait for controller closure
and delayed native audit before another Excel test. Stop and record time on
error5/mouse lock or needed approval. User authorizes closing Excel.

## 4. Evidence and traceability

`tests/integration/plan022_slice4be_inventory_uom_recording_results.md` owns exact
controller/worker IDs, timestamps, candidate boundaries and excluded attempts.

- Focused RED56 PASS/12 FAIL ->68/68 on candidate01; all prior59 checks retained.
- Guide RED73/1 ->74/74 on candidate02; exact ordered checks, complete seven-step
  and interrupted three-step observations, four missing steps, six reviewed captures.
- Candidate02 differs in only `modExpectationDraft.bas` among435 source pins;
  candidate01's focused68 is not claimed rerun on02.
- Candidate02 selections65, smoke86, Settings202 retain exact prior ordered checks;
  ten Settings captures reviewed. Normal preservation/closure and delayed audits
  pass through07:06:35.850Z; both smoke exits unassisted.
- Five compiles/cold start;36 source guards; static386 components/6722 procedures/
  145550 lines. Literal9/unresolved45/duplicate194 unchanged; all27 size caps,
  three JSON schemas, five PowerShell parses and diff checks pass.

Prior navigation/boundaries151 and Catalog/Edit50 retain catalog45 candidate scope.
Prior receipt/replay/layout evidence is not relabeled as rerun here.

## 5. Do Not Repeat

VBE normalizes argument case; case-sensitive dialog-adapter substitution silently
left the old provider active. Use verified case-insensitive replacement. Two
excluded adapter failures are not product RED. The guide-editor failure was real.
Do not kill responsive smoke finalization. Do not parse human reports or treat
modal OK/Boolean submission as proof of canonical completion.

## 6. Assumptions to Re-verify

Start a fresh bounded cursor movement/restoration proof and desktop monitor before
Excel work. Stopped `inventory-uom-desktop-20261021.json` has zero sticky failure
counters; subsequent samples were passive. Smoke's generated template is parked
byte-identically in ignored `inventory-uom-generated-template-20261021.xlsb`.
Use this PC's established push workflow and explicit `gh -R soetrain/invSys_fork`.

## 7. Open questions and blockers

Create Table/Upload and owning modal Add/Edit/Delete remain unrecorded. Worksheet
owners currently return Boolean plus human report; expose typed owner facts for
observations without changing write semantics. Preserve partial-result uncertainty
and exact contributing event IDs. The modal launcher owns mutations after form OK.

Wider D18, conformant live/full-chain fixtures, earlier RED Config evidence gaps,
comprehensive B0 and final long/multiline/provenance/incomplete/human acceptance
audit remain. Live-role fixtures still fabricate canonical rows/SYS-LIVE keys;
replace through Admin generation/owners while retaining48/32 ordered checks.

## 8. Immediate next action

Define the existing Create Table/Upload owner-result observation refinement in
Architecture/Plan022/Controls, then establish focused packaged RED before implementation.

## 9. Critical references

- Admin `frmAddInventoryItem.mBtnCreateInventoryTable_Click` /
  `mBtnUploadInventoryTable_Click`, `modAdminInventoryWorksheet`, `modAdmin.Add_InventoryItem`.
- Core `modInventoryCommands`, `modInventoryActivityCodes`, `modExpectationDraft`.
- `tests/tooling/Slice4beInventoryUomRecording.ps1`, `Slice4beInventoryGuide.ps1`,
  `Slice4beGeneralSettingsGuidePaths.ps1`, `Test-Slice4beConfigCommands.ps1`.
- Code remaining-coverage and Inventory UOM/selection/receipt integration records;
  docs CURRENT_SPEC.md, expert guidance CURRENT.md and Controls1.713.
