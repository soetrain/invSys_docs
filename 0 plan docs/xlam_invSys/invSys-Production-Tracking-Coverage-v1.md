# Production form tracking coverage audit

Version 1.4. Last reviewed: 2026-09-25 UTC, lifecycle RED and non-button source review.
Subordinate to Architecture v4.11 D18, Plan022 and `invSys-Controls-v1.md`.
This is reviewed source reachability/evidence accounting, not a new runtime
contract, catalog version, permission grant or acceptance claim. D13 behavioral
RED/GREEN does not apply to this documentation-only audit; future implementation
still requires its exact owner contract and packaged-handler RED before edits.

## Scope and reconciliation

[`frmProduction.frm`](../../../invSys_fork/src/Production/Forms/frmProduction.frm)
constructs six pages through BuildLayout and one shared Close button. It contains
82 AddButton constructions: **68 constructed buttons with Click handlers** and
**14 buttons confined to the uncalled BuildRecipeBuilderPage**. All 82 are
accounted for below. Catalog13 contains thirteen IDs mapped to
these 68 buttons: six draft controls, six lifecycle controls and UOM Retrieve. **55 constructed button
actions still require registered observation contracts and implementation.**
Registration is not full acceptance: the six draft controls pass 390 behavioral
checks but retain the visible/owner/full-chain gates recorded in the controls
catalog. Existing UOM evidence retains its recorded candidate scope.
The six lifecycle controls now pass615/615 after packaged RED316/299, retaining
the original294 and failure-extension486 identities. Owner uncertainty/pending,
tracking-off/unavailable storage, nested actions and closed-workbook guards pass
without runtime changes. Native cancellation, publication, original recording,
both Action Path methods and visible acceptance work remain open. See
[lifecycle evidence](../../../invSys_fork/tests/integration/plan022_slice4be_production_lifecycle_results.md).

The subsequent [recording gate](../../../invSys_fork/tests/integration/plan022_slice4be_production_lifecycle_paths_results.md)
exposes and protects a D5 Config-repair breach in the processor. Its intermediate
read-only correction retains24 actual actions, original journal order and real
published Activity/Designs identity through Viewer detail. The existing D18
CommandCompleted and Designs applied/awaiting mappings record668 PASS/40 expected
FAIL across708 packaged checks. Their source correction is unbuilt/not GREEN while
Windows retains the exited Excel process. Subsequent runspace recovery lets the
original controller verify settings restoration and package preservation at01:45 UTC.
Source static/layout checks pass; both presentations and full acceptance remain open. No additional
controls are registered by this checkpoint, and the55 pending buttons remain pending.

The census checks construction and handler binding, not current enabled state
or successful workflow execution. The original census performed no business
action; subsequent lifecycle evidence uses disposable packaged fixtures.
A lexical Click/Change name alone does not prove operator
reachability. Legacy construction, hidden controls and user versus programmatic
events are handled separately below. Ribbon/worksheet launchers and other forms
remain outside this form census and still require their own complete accounting.

## Constructed buttons

The exact UI name and handler identify pending entries; `Pending` means no
registered ControlId/outcome/owner contract has been established for this action.
Do not treat these UI names as new ControlIds or log every handler automatically.
All entries belong to the Production role in `invSys.Operations.xlam`.
The public launcher `mProduction.BtnOpenProductionForm` currently checks PROD_POST;
that entry check does not establish each retained form action's current-session
capability or captured-context guard. Check each owning boundary before specifying
its observation contract. In particular, do not infer UOM publication permission
from permission to stage or export a catalog worksheet.

### Form (1)

| UI control | Caption | Actual handler | Tracking ID/status |
|---|---|---|---|
| `btnProductionClose` | Close | `mBtnClose_Click` | Pending |

### Process Designer (27)

| UI control | Caption | Actual handler | Tracking ID/status |
|---|---|---|---|
| `btnProcessClear` | Clear | `mBtnProcessClear_Click` | `PRODUCTION_PROCESS_CLEAR` |
| `btnProcessInstructionAdd` | Add | `mBtnProcessInstructionAdd_Click` | Pending |
| `btnProcessInstructionDown` | Down | `mBtnProcessInstructionDown_Click` | Pending |
| `btnProcessInstructionRemove` | Remove | `mBtnProcessInstructionRemove_Click` | Pending |
| `btnProcessInstructionUp` | Up | `mBtnProcessInstructionUp_Click` | Pending |
| `btnProcessInstructionUpdate` | Update | `mBtnProcessInstructionUpdate_Click` | Pending |
| `btnProcessLoad` | View Process | `mBtnProcessLoad_Click` | Pending |
| `btnProcessNew` | New Process | `mBtnProcessNew_Click` | `PRODUCTION_PROCESS_NEW` |
| `btnProcessObsolete` | Obsolete | `mBtnProcessObsolete_Click` | `PRODUCTION_PROCESS_OBSOLETE` (partial) |
| `btnProcessOutputAdd` | Add | `mBtnProcessOutputAdd_Click` | Pending |
| `btnProcessOutputDown` | Down | `mBtnProcessOutputDown_Click` | Pending |
| `btnProcessOutputRemove` | Remove | `mBtnProcessOutputRemove_Click` | Pending |
| `btnProcessOutputUp` | Up | `mBtnProcessOutputUp_Click` | Pending |
| `btnProcessOutputUpdate` | Update | `mBtnProcessOutputUpdate_Click` | Pending |
| `btnProcessRefresh` | Refresh | `mBtnProcessRefresh_Click` | Pending |
| `btnProcessRelease` | Release | `mBtnProcessRelease_Click` | `PRODUCTION_PROCESS_RELEASE` (partial) |
| `btnProcessRequirementAdd` | Add | `mBtnProcessRequirementAdd_Click` | Pending |
| `btnProcessRequirementDown` | Down | `mBtnProcessRequirementDown_Click` | Pending |
| `btnProcessRequirementRemove` | Remove | `mBtnProcessRequirementRemove_Click` | Pending |
| `btnProcessRequirementUp` | Up | `mBtnProcessRequirementUp_Click` | Pending |
| `btnProcessRequirementUpdate` | Update | `mBtnProcessRequirementUpdate_Click` | Pending |
| `btnProcessReuse` | Edit as New Version | `mBtnProcessReuse_Click` | Pending |
| `btnProcessSave` | Save Draft | `mBtnProcessSave_Click` | `PRODUCTION_PROCESS_SAVE` (partial) |
| `btnProcessValidate` | Validate | `mBtnProcessValidate_Click` | `PRODUCTION_PROCESS_VALIDATE` |
| `btnProcessWorksheetAddAlternative` | Add Acceptable Item | `mBtnProcessWorksheetAddAlternative_Click` | Pending |
| `btnProcessWorksheetCreate` | Send Process to Sheet | `mBtnProcessWorksheetCreate_Click` | Pending |
| `btnProcessWorksheetRetrieve` | Retrieve Selected Process | `mBtnProcessWorksheetRetrieve_Click` | Pending |

### Recipe Designer (16)

| UI control | Caption | Actual handler | Tracking ID/status |
|---|---|---|---|
| `btnRecipeAddProcess` | Add Process | `mBtnRecipeAddProcess_Click` | Pending |
| `btnRecipeAutoOrder` | Auto Order | `mBtnRecipeAutoOrder_Click` | Pending |
| `btnRecipeClear` | Clear | `mBtnRecipeClear_Click` | `PRODUCTION_RECIPE_CLEAR` |
| `btnRecipeConnect` | Connect | `mBtnRecipeConnect_Click` | Pending |
| `btnRecipeDisconnect` | Disconnect | `mBtnRecipeDisconnect_Click` | Pending |
| `btnRecipeLoad` | Load | `mBtnRecipeLoad_Click` | Pending |
| `btnRecipeMoveDown` | Move Down | `mBtnRecipeMoveDown_Click` | Pending |
| `btnRecipeMoveUp` | Move Up | `mBtnRecipeMoveUp_Click` | Pending |
| `btnRecipeNew` | New Recipe | `mBtnRecipeNew_Click` | `PRODUCTION_RECIPE_NEW` |
| `btnRecipeObsolete` | Obsolete | `mBtnRecipeObsolete_Click` | `PRODUCTION_RECIPE_OBSOLETE` (partial) |
| `btnRecipeRefresh` | Refresh | `mBtnRecipeRefresh_Click` | Pending |
| `btnRecipeRelease` | Release | `mBtnRecipeRelease_Click` | `PRODUCTION_RECIPE_RELEASE` (partial) |
| `btnRecipeRemoveProcess` | Remove Process | `mBtnRecipeRemoveProcess_Click` | Pending |
| `btnRecipeSave` | Save Draft | `mBtnRecipeSave_Click` | `PRODUCTION_RECIPE_SAVE` (partial) |
| `btnRecipeUpdateConnection` | Update | `mBtnRecipeUpdateConnection_Click` | Pending |
| `btnRecipeValidate` | Validate Recipe | `mBtnRecipeValidate_Click` | `PRODUCTION_RECIPE_VALIDATE` |

### Ingredients Assignment (7)

| UI control | Caption | Actual handler | Tracking ID/status |
|---|---|---|---|
| `btnAssignAdd` | Add Acceptable | `mBtnAssignAdd_Click` | Pending |
| `btnAssignClear` | Clear | `mBtnAssignClear_Click` | Pending |
| `btnAssignIngredient` | Select Requirement | `mBtnAssignIngredient_Click` | Pending |
| `btnAssignRecipe` | Select Process | `mBtnAssignRecipe_Click` | Pending |
| `btnAssignRefresh` | Refresh | `mBtnAssignRefresh_Click` | Pending |
| `btnAssignRemove` | Remove Row | `mBtnAssignRemove_Click` | Pending |
| `btnAssignSave` | Save Alternatives | `mBtnAssignSave_Click` | Pending |

### Production Run - List (10)

| UI control | Caption | Actual handler | Tracking ID/status |
|---|---|---|---|
| `btnApplyBatchScale` | Apply Scale | `mBtnApplyBatchScale_Click` | Pending |
| `btnLoaderClear` | Clear Run | `mBtnLoaderClear_Click` | Pending |
| `btnLoaderLoad` | Load Recipe | `mBtnLoaderLoad_Click` | Pending |
| `btnLoaderRefresh` | Refresh | `mBtnLoaderRefresh_Click` | Pending |
| `btnManagerApplyOutput` | Complete Run | `mBtnManagerApplyOutput_Click` | Pending |
| `btnManagerCheckIn` | Check In | `mBtnManagerCheckIn_Click` | Pending |
| `btnManagerNext` | Next Batch | `mBtnManagerNext_Click` | Pending |
| `btnManagerPrint` | Print Recall | `mBtnManagerPrint_Click` | Pending |
| `btnManagerRefresh` | Refresh | `mBtnManagerRefresh_Click` | Pending |
| `btnRunApplyPalette` | Apply | `mBtnRunApplyPalette_Click` | Pending |

### Production Run - Tree (3)

| UI control | Caption | Actual handler | Tracking ID/status |
|---|---|---|---|
| `btnRunTreeApplyPalette` | Apply | `mBtnRunTreeApplyPalette_Click` | Pending |
| `btnRunTreeCollapseAll` | Collapse | `mBtnRunTreeCollapseAll_Click` | Pending |
| `btnRunTreeExpandAll` | Expand | `mBtnRunTreeExpandAll_Click` | Pending |

### Production Settings (4)

| UI control | Caption | Actual handler | Tracking ID/status |
|---|---|---|---|
| `btnOutputRegulationApply` | Apply Regulation | `mBtnOutputRegulationApply_Click` | Pending |
| `btnOutputRegulationClear` | Clear Override | `mBtnOutputRegulationClear_Click` | Pending |
| `btnUomCatalogRetrieve` | Retrieve UOM Catalog | `mBtnUomCatalogRetrieve_Click` | `PRODUCTION_UOM_RETRIEVE` |
| `btnUomCatalogSend` | Edit UOM Catalog on Sheet | `mBtnUomCatalogSend_Click` | Pending |

## Unconstructed Recipe Builder controls

BuildRecipeBuilderPage has no invocation in source; BuildLayout constructs the newer Process and Recipe designers instead. The following 14 buttons are excluded from the current operator-button count on that evidence, not retired or deleted. Existing test helpers and legacy service paths are not permission to expose them or fall back to legacy storage when DesignsEnabled=True.

| UI control | Caption | Unreachable form event |
|---|---|---|
| `btnBuilderRefresh` | Refresh | `mBtnBuilderRefresh_Click` |
| `btnBuilderNew` | New Recipe | `mBtnBuilderNew_Click` |
| `btnBuilderLoad` | Load Selected | `mBtnBuilderLoad_Click` |
| `btnBuilderSave` | Save Recipe | `mBtnBuilderSave_Click` |
| `btnBuilderProcess` | Add Process Table | `mBtnBuilderProcess_Click` |
| `btnBuilderFormulas` | Save Formulas | `mBtnBuilderFormulas_Click` |
| `btnBuilderClear` | Clear Builder | `mBtnBuilderClear_Click` |
| `btnBuilderRelease` | Release for Production | `mBtnBuilderRelease_Click` |
| `btnLineUomAdd` | Add UOM | `mBtnLineUomAdd_Click` |
| `btnLineAdd` | Add Line | `mBtnLineAdd_Click` |
| `btnLineUpdate` | Update Line | `mBtnLineUpdate_Click` |
| `btnLineRemove` | Remove Line | `mBtnLineRemove_Click` |
| `btnLineMoveUp` | Move Up | `mBtnLineMoveUp_Click` |
| `btnLineMoveDown` | Move Down | `mBtnLineMoveDown_Click` |

## Non-button event audit

The source contains 34 non-button Click/Change handlers. Three belong to the
unconstructed Recipe Builder and one to the hidden internal connection list.
The remaining30 are reviewed below against their current helpers. These are source
facts and required test distinctions, not new IDs, approved terminal outcomes or
runtime acceptance. Deliberate selections still need observation contracts; text
changes and mirrored programmatic events must not become keystroke surveillance
or duplicate actions under D18.

| Handler | Current source effect / required distinction |
|---|---|
| `mLstProcesses_Click` | Reports selected identity only; does not load or persist a draft |
| `mLstProcessRequirements_Click` | Copies selected row into editor and applies quantity mode; programmatic combo changes can cascade |
| `mLstProcessOutputs_Click` | Copies selected row into editor, refreshes UOM choices and applies quantity mode; suppress cascaded observations |
| `mCmbRequirementQtyMode_Change` | Locks/unlocks editor fields; ACTUAL clears three entered fields in ApplyRequirementQtyMode; not a display-only selection |
| `mCmbProcessOutputQtyMode_Change` | Locks/unlocks editor fields; ACTUAL clears three entered fields in ApplyOutputQtyMode; not a display-only selection |
| `mCmbOutputRegulationScope_Change` | Rebuilds regulation/node choices and clears regulation editor; Recipe scope reads selected Process version |
| `mCmbOutputRegulationNode_Change` | Same rebuild only for Recipe scope; mLoading suppresses internal node-list changes |
| `mLstOutputRegulations_Click` | Copies regulation flag/floor/ceiling into editor; does not invoke Apply |
| `mLstProcessInstructions_Click` | Copies selected instruction into editor; no mLoading guard in this handler, so deliberate input needs separate proof |
| `mCmbConnectionFromNode_Change` | Refreshes available outputs and dependent choices; distinguish initiating selection from helper cascades |
| `mCmbConnectionOutput_Change` | Refreshes compatible downstream choices; distinguish initiating selection from helper cascades |
| `mCmbConnectionToNode_Change` | Binds matching Process requirement defaults (quantity/percent/UOM); does not save the connection |
| `mLstRecipeConnections_Click` | Hidden internal list; visible owner is mLstRecipeConnectionDisplay_Click |
| `mLstRecipeConnectionDisplay_Click` | Existing row selects hidden list and loads editor; finished-output row prepares routing choices. One visible action must not duplicate hidden-list/combo callbacks |
| `mLstBuilderRecipes_Click` | Unconstructed Recipe Builder |
| `mLstBuilderLines_Click` | Unconstructed Recipe Builder |
| `mCmbLineIo_Change` | Unconstructed Recipe Builder |
| `mLstAssignRecipes_Click` | Reads selected Process, replaces local requirements/alternatives and refreshes allowed items; same owner helper is used by Select Process button |
| `mLstAssignIngredients_Click` | Refreshes allowed items for selected requirement; same owner helper is used by Select Requirement button |
| `mTxtInventorySearch_Change` | Filters inventory through RefreshInventoryList; no entered search text or per-keystroke observations |
| `mCmbRunLocation_Change` | Mirrors Tree selection under mLoading; reusable branch refreshes controls, non-reusable branch can clear incompatible staging allocations |
| `mCmbTreeRunLocation_Change` | Mirrors List selection under mLoading with the same two branches; test one deliberate action, not two observations |
| `mTxtPaletteSplit_Change` | Computes mirrored quantity and allocation visual state; mUpdatingPaletteInputs prevents recursive edits; does not invoke Apply Split |
| `mTxtTreePaletteSplit_Change` | Tree counterpart of split-input computation; no per-keystroke observation |
| `mTxtPaletteQty_Change` | Computes mirrored percentage and allocation visual state; does not invoke Apply Split |
| `mTxtTreePaletteQty_Change` | Tree counterpart of quantity-input computation; no per-keystroke observation |
| `mTxtOutputReal_Change` | Loaded reusable run stages or clears actual-output quantity in memory through StageSelectedReusableActualOutput(False); not merely editor formatting |
| `mTxtRunBatchNote_Change` | Loaded reusable run calls SetReusableRunBatchNote; length/frozen-note refusal can occur before any inventory submission |
| `mLstLoaderLines_Click` | Reusable branch displays selected line; non-reusable branch refreshes acceptable inventory palette |
| `mLstRunPalette_Click` | Loads selected allocation editor; guard programmatic selection and linked input changes |
| `mLstRunTree_Click` | Parent toggles expansion/rebuilds Tree; leaf loads allocation editor. Distinguish navigation outcomes before assigning IDs |
| `mLstManagerOutput_Click` | Loads selected output editor; reusable helper changes Actual Output text under mLoading to avoid staging it again |
| `mCmbRunProcess_Change` | Mirrors Tree filter under mLoading, then refreshes reusable palette/instructions or non-reusable palette; no run submission |
| `mCmbTreeRunProcess_Change` | Mirrors List filter under mLoading with same refresh owner; test one deliberate action |

Specific protecting cases still required: quantity-mode selection must preserve
the existing editor-clearing behavior without reporting persistence; a displayed
connection selection must not double count its hidden list or combo cascades;
List/Tree process and location mirrors must observe only the initiating action;
Tree parent expansion and leaf selection must remain distinguishable. Actual
Output and Batch Note update reusable in-memory state during Change, so an eventual
semantic commit contract must account for that state without capturing entered
values or emitting one event per keystroke. Source review does not choose a new
commit boundary or authorize changing these behaviors.

The non-reusable location branch calls ClearMismatchedRunLocationAllocations,
then ClearRunAllocationForListRow, which can clear SPLIT %/QUANTITY cells in a
captured staging table as well as local List/Tree state. Therefore a generic
read-only Navigation label would be misleading without branch-specific owner
facts. This source observation does not accept that branch, authorize fallback
when DesignsEnabled=True, or change immutable System_Key identity requirements.

There is no mPages_Change handler for the six visible tabs. Deliberate page
selection therefore still needs its own observation contract. Selectable controls
without a dedicated current handler also need review (including saved Recipes,
released Processes, Recipe nodes and staging selections). This handler list is
not a substitute for the complete control inventory in the main catalog.
The target-output scaling checkbox/combo/quantity are explicitly disabled in
BuildLoaderPage and labeled as a later solver; they are not active commands.
Run - Tree is experimental but constructed; its three buttons are pending,
not silently excluded because Run - List is the Release 1 proving path.

## Owner boundaries that must govern the next contracts

1. Process/Recipe Save, Release and Obsolete now call ExecuteDesignerLifecycle,
   SubmitDesignerAction and modProductionReusableDesigns.SubmitReusableDesignEvent.
   The six registered handlers now carry cProductionLifecycleFacts; other callers
   of SubmitDesignerAction are not thereby covered by those six control contracts.
   QueueDesignEventCurrent already returns the exact event ID and queue Boolean.
   The helper subsequently runs a processor batch and compares the definition's
   projected status. Its final Boolean can be False after successful queueing;
   a True Boolean and global applied-count are not exact application evidence for
   that submitted event. Preserve those distinct facts through typed results.
   Never parse EventID out of the operator report or infer per-event application
   from a matching status that may already have existed. Retain correlation even
   when processing/projection verification fails after queue acceptance.
2. RefreshReusableDesignLists reads four Process/Recipe projections and fills
   lists. The current handlers unconditionally display a refreshed status after
   it returns. A future REFRESHED observation must come from the owning read's
   supported result contract; status text alone is not proof that every source
   was available. Do not add repair/publication or canonical authority reads to
   Viewer to compensate for missing evidence.
3. Requirement/output/instruction edits, Recipe nodes/connections, output
   regulation and alternatives have distinct local-draft and persistence
   boundaries. Map actions to those actual boundaries before assigning outcomes;
   names such as Apply, Save or Clear do not determine Domain effect.
4. UOM Send calls modProductionUomCatalog.SendUomCatalogToWorksheet (source file
   modProductionUomCatalogWorksheet.bas): it writes the captured staging sheet.
   It unlists an existing staging table and clears that worksheet before writing
   the catalog. User-column preservation for that path is not proved by this
   audit; protect it with a focused behavioral test before changing the writer.
   It is not a read-only observation and does not publish warehouse Config.
   Retrieve uses modProductionUomAction and the existing authorized Config owner.
   Preserve unknown user columns and unrelated user content in any future change.
5. Run selection, allocation, scale, actual-output and note helpers can react to
   programmatic changes and synchronize List/Tree controls. Observe deliberate
   operator actions once. Check In and Complete Run need exact inventory-event
   correlation from their owner. Next Batch and Print Recall require their own
   supported results; generic handler return must not assert persistence.
6. Close and public/repeated launcher behavior must preserve captured-workbook
   reuse and invalidate observation/recording context correctly. A visible control
   or an enabled legacy helper does not authorize a different session/warehouse.

## Validation and next implementation grouping

The source audit reconciles every AddButton construction, all six BuildLayout
page calls, every bound button's actual Click procedure and the sole uncalled
Recipe Builder constructor. The thirteen registered mappings are cross-checked
against DesignerDraftAction/ExecuteDesignerLifecycle and modProductionUomAction;
all34 non-button handlers and the relevant mutation/mirroring helpers are reviewed.
No runtime
tracking or catalog mutation occurs. Private source-only census/verification
artifacts are ignored; this reviewed record contains no operational values.

Use the existing registered-control candidates to finish their remaining visible and
regression gates once desktop input recovers. For broader implementation, review
the grouped local-edit/selection, lifecycle/queue, worksheet/Config, run/inventory
and launcher/context owners from this census; specify complete owner facts and
focused packaged tests before implementing each group. Do not substitute the
thirteen registered Production IDs for comprehensive Operations/Admin coverage.
