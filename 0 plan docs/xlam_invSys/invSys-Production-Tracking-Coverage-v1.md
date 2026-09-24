# Production form tracking coverage audit

Version 1.0. Last reviewed: 2026-09-24 UTC, code `9cf03db`.
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
accounted for below. The existing activity catalog contains seven IDs mapped to
these 68 buttons: the six draft controls and UOM Retrieve. **61 constructed button
actions still require registered observation contracts and implementation.**
Registration is not full acceptance: the six draft controls pass 390 behavioral
checks but retain the visible/owner/full-chain gates recorded in the controls
catalog. Existing UOM evidence retains its recorded candidate scope.

The census checks construction and handler binding, not current enabled state
or successful workflow execution. No business action, workbook, add-in or printer
was opened or executed. A lexical Click/Change name alone does not prove operator
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
| `btnProcessObsolete` | Obsolete | `mBtnProcessObsolete_Click` | Pending |
| `btnProcessOutputAdd` | Add | `mBtnProcessOutputAdd_Click` | Pending |
| `btnProcessOutputDown` | Down | `mBtnProcessOutputDown_Click` | Pending |
| `btnProcessOutputRemove` | Remove | `mBtnProcessOutputRemove_Click` | Pending |
| `btnProcessOutputUp` | Up | `mBtnProcessOutputUp_Click` | Pending |
| `btnProcessOutputUpdate` | Update | `mBtnProcessOutputUpdate_Click` | Pending |
| `btnProcessRefresh` | Refresh | `mBtnProcessRefresh_Click` | Pending |
| `btnProcessRelease` | Release | `mBtnProcessRelease_Click` | Pending |
| `btnProcessRequirementAdd` | Add | `mBtnProcessRequirementAdd_Click` | Pending |
| `btnProcessRequirementDown` | Down | `mBtnProcessRequirementDown_Click` | Pending |
| `btnProcessRequirementRemove` | Remove | `mBtnProcessRequirementRemove_Click` | Pending |
| `btnProcessRequirementUp` | Up | `mBtnProcessRequirementUp_Click` | Pending |
| `btnProcessRequirementUpdate` | Update | `mBtnProcessRequirementUpdate_Click` | Pending |
| `btnProcessReuse` | Edit as New Version | `mBtnProcessReuse_Click` | Pending |
| `btnProcessSave` | Save Draft | `mBtnProcessSave_Click` | Pending |
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
| `btnRecipeObsolete` | Obsolete | `mBtnRecipeObsolete_Click` | Pending |
| `btnRecipeRefresh` | Refresh | `mBtnRecipeRefresh_Click` | Pending |
| `btnRecipeRelease` | Release | `mBtnRecipeRelease_Click` | Pending |
| `btnRecipeRemoveProcess` | Remove Process | `mBtnRecipeRemoveProcess_Click` | Pending |
| `btnRecipeSave` | Save Draft | `mBtnRecipeSave_Click` | Pending |
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

The source contains 34 non-button Click/Change handlers. Three belong to the unconstructed Recipe Builder and one to the hidden internal connection list. The remaining 30 require semantic review: deliberate selections may need observation; text changes and mirrored programmatic events must not become keystroke surveillance or duplicate actions. This list does not invent terminal outcomes.

| Handler | Source reachability / next review |
|---|---|
| `mLstProcesses_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstProcessRequirements_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstProcessOutputs_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbRequirementQtyMode_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbProcessOutputQtyMode_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbOutputRegulationScope_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbOutputRegulationNode_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstOutputRegulations_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstProcessInstructions_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbConnectionFromNode_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbConnectionOutput_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbConnectionToNode_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstRecipeConnections_Click` | Hidden internal list; visible owner is mLstRecipeConnectionDisplay_Click |
| `mLstRecipeConnectionDisplay_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstBuilderRecipes_Click` | Unconstructed Recipe Builder |
| `mLstBuilderLines_Click` | Unconstructed Recipe Builder |
| `mCmbLineIo_Change` | Unconstructed Recipe Builder |
| `mLstAssignRecipes_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstAssignIngredients_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mTxtInventorySearch_Change` | Input-change helper; review semantic commit boundary, never capture entered text |
| `mCmbRunLocation_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbTreeRunLocation_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mTxtPaletteSplit_Change` | Input-change helper; review semantic commit boundary, never capture entered text |
| `mTxtTreePaletteSplit_Change` | Input-change helper; review semantic commit boundary, never capture entered text |
| `mTxtPaletteQty_Change` | Input-change helper; review semantic commit boundary, never capture entered text |
| `mTxtTreePaletteQty_Change` | Input-change helper; review semantic commit boundary, never capture entered text |
| `mTxtOutputReal_Change` | Input-change helper; review semantic commit boundary, never capture entered text |
| `mTxtRunBatchNote_Change` | Input-change helper; review semantic commit boundary, never capture entered text |
| `mLstLoaderLines_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstRunPalette_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstRunTree_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mLstManagerOutput_Click` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbRunProcess_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |
| `mCmbTreeRunProcess_Change` | Constructed control; distinguish deliberate selection from mLoading/programmatic updates |

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

1. Process/Recipe Save, Release and Obsolete call SubmitProcessAction or
   SubmitRecipeAction, then modProductionReusableDesigns.SubmitReusableDesignEvent.
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
Recipe Builder constructor. The seven known mappings are cross-checked against
the actual DesignerDraftAction and modProductionUomAction call paths; no runtime
tracking or catalog mutation occurs. Private source-only census/verification
artifacts are ignored; this reviewed record contains no operational values.

Use the existing six-control candidate to finish its remaining visible and
regression gates once desktop input recovers. For broader implementation, review
the grouped local-edit/selection, lifecycle/queue, worksheet/Config, run/inventory
and launcher/context owners from this census; specify complete owner facts and
focused packaged tests before implementing each group. Do not substitute the
seven registered Production IDs for comprehensive Operations/Admin coverage.
