# invSys Form Controls v1

**Version:** 1.0

**Inventory date:** 2026-08-04

**Architecture:** invSys v4.11, Release 1
**Scope:** every checked-in VBA UserForm source file used by Core, Admin,
Receiving, Production, or Shipping, including active controls, runtime-created
controls, inactive source shells, status surfaces, and generated column labels.

## 1. Purpose and authority

This is a readable control catalog. It describes the current implementation;
it does not make every listed control architecturally valid or release-ready.
The normative authority remains `invSys-Design-v4.11.md`.

The catalog was derived from:

- all 34 checked-in `.frm` files under `src`;
- all four checked-in `.frx` designer blobs;
- the current deployed five-package set in `deploy/current`;
- runtime control-construction calls and their event handlers; and
- the 2026-08-04 operator screenshots and acceptance report.

The 34 source files produce 31 distinct deployed form components. The three
role copies named `ufDynItemSearchTemplate` collide by component name and are
not imported into the consolidated Operations project. Current search windows
use the role-named empty shells plus `Core.cDynItemSearch`, which adds their
controls at runtime.

This document covers forms only. Ribbon controls and worksheet buttons are
outside its scope. The older `controls.md` remains a broader, historical
status list; this file supersedes it only for the form-control inventory.

### Status terms

| Term | Meaning |
|---|---|
| Active | A current launcher or workflow constructs and uses the form. |
| Runtime-generated | The `.frm` has an empty designer and creates controls in VBA. |
| Designer-backed | Controls are stored in a checked-in `.frx` blob. |
| Shell | The component is packaged but has no deployed controls or useful current workflow. |
| Deprecated | Source explicitly identifies the form as superseded. |

## 2. Acceptance findings that affect this catalog

### Seed Demo Inventory: automated round trip GREEN; visible retest pending

`frmSeedInventory` displayed and its public callback reported one applied
processor event. The operator then refreshed, but the three demo entities were
not visible in that checkpoint. Subsequent read-only inspection found the demo
entities in all three saved role projections, and the expanded packaged test
proved the same three unique `System_Key` values with `Condition=GOOD` through
canonical inventory, snapshot, saved Receiving workbook, and the Receiving
form's actual Refresh handler. The operator checkpoint remains failed until the
rows are verified in the top Inventory list on the Receiving tab; Recipe
Builder and Shipping's shippable list are not raw inventory views.

### Production resizing: corrected; visible retest pending

`frmProduction` maximized its native window while the MultiPage and child
controls remained near their base dimensions in the upper-left corner. The
focused validator reproduced the failure at DPI-derived `Zoom=60`. The deployed
correction holds the form at `Zoom=100`; packaged validation now passes native
minimize/restore/maximize/restore and bounds/fill/overlap checks for all four
pages. A visible maximize/restore retest remains required.

### Shipping still exposes prohibited `ROW` controls

`frmShipmentsTally` currently creates `lblRow`, `txtRow`, `hdrShipRow`, two
generated `hdrLineRow*` labels, and an unused `hdrReadyRow` builder. D14
prohibits `ROW` as a managed runtime header, display key, compatibility field,
or authority path. These controls are listed because they exist, and are
explicitly marked **architecturally invalid**. They must become
`System_Key`-based controls under a focused D13 correction.

## 3. Complete form inventory

| Package/source area | Form | Current status | Control surface |
|---|---|---|---|
| Admin | `frmAddInventoryItem` | Active, runtime-generated | Add/edit inventory metadata and custom fields |
| Admin | `frmAdminControls` | Shell | No deployed controls; two handlers reference absent legacy buttons |
| Admin | `frmAdminDesignLifecycle` | Active, runtime-generated | Design refresh/import/release/obsolete actions |
| Admin | `frmAdminEmail` | Shell | No deployed controls or form code |
| Admin | `frmAdminSettings` | Active, runtime-generated | Config, connection policy, carriers, UOM catalog |
| Admin | `frmCreateDeleteUser` | Active, runtime-generated | Warehouse user and capability maintenance |
| Admin | `frmCreateWarehouse` | Active, designer-backed plus runtime buttons | Greenfield warehouse creation |
| Admin | `frmEditUser` | Shell | No deployed controls; legacy event code references absent controls |
| Admin | `frmReAuthGate` | Active, designer-backed | Destructive-action re-authentication |
| Admin | `frmRetireMigrateWarehouse` | Active, designer-backed plus runtime controls | Archive, retire, migrate, or delete warehouse |
| Admin | `frmSeedInventory` | Active, runtime-generated | Choose target and seed demo inventory |
| Admin | `frmSetupTesterStation` | Active, reused designer plus runtime controls | Disposable tester-station setup and cleanup |
| Admin | `ufAdminItemSearch` | Empty shell | No current caller or deployed controls |
| Admin | `ufDynItemSearchTemplate` | Empty shell | Packaged template name; no controls by itself |
| Core | `frmItemSearch` | Deprecated/incomplete | Code expects `txtBox` and `lstBox`, but current deployed form has no controls |
| Core | `frmSignIn` | Active, runtime-generated | invSys user authentication |
| Core | `frmWarehouseConnection` | Active, runtime-generated | NAS/root connection and warehouse target selection |
| Production | `frmCreateRecipeTable` | Shell | No deployed controls |
| Production | `frmCreateSubstitutionList` | Shell | No deployed controls |
| Production | `frmIngredientPalette` | Shell | No deployed controls |
| Production | `frmProduction` | Active, runtime-generated | Recipe, assignment, run-list, and run-tree workflows |
| Production | `frmSubstitution` | Shell | No deployed controls |
| Production | `ufProductionItemSearch` | Runtime template | Four controls added by `cDynItemSearch` when invoked |
| Production | `ufDynItemSearchTemplate` | Duplicate source shell | Not imported into consolidated Operations package |
| Receiving | `frmReceiving` | Active, runtime-generated | Receiving and non-operational Purchasing tab |
| Receiving | `frmReceivingSavedList` | Shell | No deployed controls |
| Receiving | `ufReceivingItemSearch` | Empty role shell | No current caller; no controls added in the current path |
| Receiving | `ufDynItemSearchTemplate` | Duplicate source shell | Not imported into consolidated Operations package |
| Shipping | `frmBoxVersionSaveChoice` | Active, runtime-generated | Choose update-versus-new version behavior |
| Shipping | `frmShipmentsTally` | Active, runtime-generated | Shipping, Box Builder, and Box Maker tabs |
| Shipping | `frmShippingCreateList` | Shell | No deployed controls |
| Shipping | `frmShippingSavedList` | Shell | No deployed controls |
| Shipping | `ufShippingItemSearch` | Runtime template | Four controls added by `cDynItemSearch` when invoked |
| Shipping | `ufDynItemSearchTemplate` | Duplicate source shell | Not imported into consolidated Operations package |

## 4. Admin forms

### 4.1 `frmAddInventoryItem` — Add/Edit Inventory Item

This form edits inventory item metadata and queues the supported add/edit
workflow. It does not make a worksheet row the durable identity.

| Control | Type / displayed text | Purpose |
|---|---|---|
| `btnAddMode` | Button — **Add Item** | Selects add mode and clears add fields. |
| `btnEditMode` | Button — **Edit Item** | Selects edit mode and exposes inventory search/results. |
| `cmbEditItem` | Combo box — Inventory item | Accepts or selects the item to edit. |
| `lstEditItemResults` | List box | Shows candidate inventory items for edit selection. |
| `txtItemName` | Text box — Item name * | Required descriptive item name. |
| `cmbUom` | Combo box — UOM * | Selects the unit of measure. |
| `txtQty` | Combo box — Starting qty * | Accepts the initial quantity in add mode. |
| `txtLocation` | Text box — Default location | Sets the initial/default location attribute. |
| `txtCategory` | Text box — Category | Stores category metadata. |
| `txtDescription` | Text box — Description | Stores the long description. |
| `txtVendorName` | Text box — Vendor(s) | Stores vendor display metadata. |
| `txtVendorCode` | Text box — Vendor code | Stores the vendor-specific code. |
| `txtExternalCode` | Text box — External code | Stores a non-invSys external code. |
| `txtImagePath` | Text box — Picture path/URL | Stores an image reference. |
| `txtEditReason` | Text box — Why the edit? * | Requires an audit reason in edit mode. |
| `txtCustomName` | Text box — Additional field | Names a proposed custom field. |
| `txtCustomValue` | Text box — Value | Supplies the proposed custom-field value. |
| `btnAddField` | Button — **Add** | Adds the name/value pair to the pending custom-field list. |
| `lstCustomFields` | List box | Displays pending custom fields. |
| `btnRemoveField` | Button — **Remove** | Removes the selected pending custom field. |
| `btnOK` | Button — **Add Item** or edit action | Validates and submits the current mode. |
| `btnCancel` | Button — **Cancel** | Closes without submitting. |
| `lblTitle` | Label | Shows the current add/edit title. |
| `lblContext` | Label | Explains that invSys generates the internal item code. |
| `lblGenerated` | Label | Shows generated identifier/result context. |
| `lblStatus` | Label | Displays validation and completion status. |
| `lblEditItem`, `lblItemName`, `lblUom`, `lblQty`, `lblLocation`, `lblCategory`, `lblDescription`, `lblVendorName`, `lblVendorCode`, `lblExternalCode`, `lblImagePath`, `lblEditReason`, `lblCustomName`, `lblCustomValue` | Field labels | Identify the controls listed above. |

### 4.2 `frmAdminDesignLifecycle` — Designs Lifecycle

| Control | Type / displayed text | Purpose |
|---|---|---|
| generated label | **Design versions** | Section heading. |
| `lstDesigns` | Multi-column list box | Lists design/version/status records. |
| `btnRefresh` | Button — **Refresh** | Reloads the design list. |
| generated label | **Audit note** | Identifies the required note field. |
| `txtNote` | Multiline text box | Captures the release/obsolete audit note. |
| `btnRelease` | Button — **Release selected** | Requests release of the selected design version. |
| `btnObsolete` | Button — **Obsolete selected** | Requests obsolescence of the selected version. |
| `btnMigrate` | Button — **Import Legacy Recipes** | Invokes the explicit recipe-migration action. This is not old inventory migration. |
| `lblStatus` | Label | Displays action results or errors. |
| `btnClose` | Button — **Close** | Closes the form. |

### 4.3 `frmAdminSettings` — Warehouse Settings

| Area | Controls | Purpose |
|---|---|---|
| Heading/context | `lblTitle`, `lblConfigWorkbook`, `lblStatus` | Shows **Warehouse Settings**, the current config workbook, and action status. |
| Config grid | `lblConfigKeyHeader`, `lblConfigValueHeader`, `lblConfigTypeHeader`, `lblConfigScopeHeader`, `lblConfigRequiredHeader`, `lstConfig` | Lists config key, value, type, scope, and required status. |
| Selected config | `lblSelectedKey`, `txtConfigKey`, `lblSelectedValue`, `txtConfigValue` | Shows the selected key and permits value editing. |
| Config actions | `btnSaveConfig` (**Save Value**), `btnReloadConfig` (**Reload**) | Saves the selected value or reloads config. |
| Connection policy | `lblServerConnection`, `chkManualServerCredentials`, `lblServerConnectionScope`, `btnSaveConnectionPolicy` | Controls whether Connect Server always requests manual credentials for this Windows user. |
| Carrier catalog | `lblSection`, `lblCarrier`, `txtCarrier`, `btnAdd`, `btnRemove`, `btnReset`, `lstCarriers` | Adds, removes, resets, and lists Shipping carriers. |
| UOM catalog | `lblUomSection`, `lblUom`, `txtUom`, `btnUomAdd`, `btnUomRemove`, `btnUomReset`, `lstUoms` | Adds, removes, resets, and lists recipe UOM values. |
| Exit | `btnClose` — **Close** | Closes the settings form. |

### 4.4 `frmCreateDeleteUser` — Users & Roles

| Area | Controls | Purpose |
|---|---|---|
| Title/status | `lblTitle`, `lblStatus` | Shows **Users & Roles** and workflow status. |
| Warehouse root | `lblRoot`, `txtWarehouseRoot`, `btnRootFind`, `btnRootScan` | Selects or scans the warehouse root. |
| NAS credentials | `lblNasUser`, `txtNasUser`, `lblNasPassword`, `txtNasPassword`, `btnNasConnect`, `btnNasLogout` | Connects or disconnects this Windows session from the NAS. Password text is masked. |
| Warehouse selection | `lblWarehouse`, `cmbWarehouse`, `lblAuthPath`, `txtAuthPath`, `lblWarehouses`, `lstWarehouses` | Selects a discovered warehouse and shows its Auth workbook path. |
| User list | `lblUsers`, `lstUsers`, `btnRefreshUsers` | Lists and reloads users in the selected Auth workbook. |
| User identity | `lblUserId`, `txtUserId`, `lblDisplay`, `txtDisplayName` | Edits user ID and display name. |
| Credential | `lblPin`, `txtPin`, `btnGeneratePin`, `btnCopyPin` | Accepts or generates a PIN/password and copies the generated value when allowed. |
| Scope | `lblWhScope`, `txtWarehouseId`, `lblStationScope`, `txtStationId`, `chkRoamingStation` | Sets warehouse and station scope; Roaming removes the fixed-station requirement. |
| Capabilities | `lblRoles`, `chkAdmin`, `chkReceivePost`, `chkReceiveView`, `chkShipPost`, `chkProdPost`, `chkInboxProcess` | Assigns Admin maintenance, Receiving post/view, Shipping post, Production post, and inbox-processor capabilities. |
| Record actions | `btnClear`, `btnSave`, `btnDeactivate`, `btnDelete` | Clears the editor, creates/updates a user, deactivates the selected user, or deletes it through the guarded action. |
| Exit | `btnClose` | Closes the form. |

### 4.5 `frmCreateWarehouse` — Create Warehouse

Designer-backed controls:

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblIntro` | Intro label | Explains greenfield warehouse creation and initial publication. |
| `txtWarehouseId` | Text box — Warehouse ID | Required stable warehouse identifier. |
| `txtWarehouseName` | Text box — Warehouse Name | Required display name. |
| `txtStationId` | Text box — Station ID | Creates the initial station, default `S1`. |
| `txtAdminUser` | Text box — Admin User | Creates/scopes the initial admin user. |
| `txtPathLocal` | Text box — Local Path | Target warehouse hub/runtime path. |
| `txtPathSharePoint` | Text box — SharePoint Path | Locally synced invSys SharePoint root. |
| `chkPublishInitial` | Check box — Publish initial artifacts to SharePoint | Enables initial discovery/artifact publication. |
| `lblSummary` | Wrapped status label | Shows validation, bootstrap, publication, and generated-workbook results. |
| `btnOK` | Button — **Create**, **Retry Publish**, or **Close** | Advances creation/retry/completion according to form state. |
| `btnCancel` | Button — **Cancel** or **Close** | Cancels before creation or closes after completion. |
| `lblWarehouseIdError`, `lblWarehouseNameError`, `lblStationIdError`, `lblAdminUserError`, `lblPathLocalError`, `lblPathSharePointError` | Inline error labels | Show validation errors beside their fields. |
| `lblWarehouseIdCaption`, `lblWarehouseNameCaption`, `lblStationIdCaption`, `lblAdminUserCaption`, `lblPathLocalCaption`, `lblPathSharePointCaption` | Field labels | Identify the six input fields. |

Runtime-added controls:

| Control | Displayed text | Purpose |
|---|---|---|
| `btnHubPathHelperRuntime` | **Browse...** | Browses for the NAS hub/runtime folder. |
| `btnSharePointHelperRuntime` | **Find...** | Detects or browses for the locally synced SharePoint root. |
| `btnOpenReceivingRuntime` | **Open Workbook** | Opens the generated station-local Receiving workbook after successful creation. |

### 4.6 `frmReAuthGate` — Re-Authenticate

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblPrompt` | Label | Explains that credentials must be re-entered. |
| `lblAdminUserCaption`, `lblAdminUserValue` | Labels — Admin user | Identify and show the user being re-authenticated. |
| `lblRoleCaption`, `lblRoleValue` | Labels — Role | Identify and show the required role/capability context. |
| `lblPasswordCaption`, `txtPassword` | Label and masked text box — Password | Accept the re-authentication secret. |
| `lblError` | Error label | Displays failed validation/authentication. |
| `btnOK` | Button — **OK** | Attempts re-authentication. |
| `btnCancel` | Button — **Cancel** | Cancels the guarded action. |

### 4.7 `frmRetireMigrateWarehouse` — Retire / Migrate Warehouse

| Area | Controls | Purpose |
|---|---|---|
| Title/instructions | `lblTitle`, `lblSelectionIntro` | Explains source warehouse, operation mode, and archive settings. |
| Runtime root | `lblWarehouseRootRuntime`, `txtWarehouseRootPathRuntime`, `btnWarehouseRootBrowseRuntime`, `btnWarehouseRootRefreshRuntime`, `lblWarehouseRootErrorRuntime` | Selects/scans the NAS hub or warehouse runtime root and reports errors. |
| NAS credentials | `lblNasUserRuntime`, `txtNasUserRuntime`, `lblNasPasswordRuntime`, `txtNasPasswordRuntime`, `btnNasConnectRuntime` | Connects the current Windows session to the selected NAS root. |
| Discovered warehouses | `lblFoundWarehousesRuntime`, `lstSourceWarehousesRuntime` | Lists WarehouseId and folder for discovered runtimes. |
| Source/target | `lblSourceWarehouse`, `cmbSourceWarehouse`, `lblSourceWarehouseError`, `lblTargetWarehouse`, `cmbTargetWarehouse`, `lblTargetWarehouseError` | Chooses the source warehouse and optional migration target. |
| Mode frame | `fraMode` | Contains the operation-mode options. |
| Active mode options | `optArchiveOnlyRuntime`, `optArchiveRetireRuntime`, `optArchiveMigrateRuntime`, `optArchiveRetireDeleteRuntime` | Select archive-only, retire/archive, migrate inventory, or irreversible retire/archive/delete behavior. |
| Hidden designer mode options | `optArchiveOnly`, `optArchiveMigrate`, `optArchiveRetire`, `optArchiveRetireDelete` | Legacy designer controls hidden in favor of the runtime option controls. |
| Archive destination | `lblArchiveDestPath`, `txtArchiveDestPath`, `btnArchiveDestBrowseRuntime`, `lblArchiveDestPathError` | Selects and validates the output archive path. |
| Publication/re-auth | `chkPublishTombstone`, `lblReAuthError`, `lblDeleteWarning` | Enables tombstone publication and shows re-auth/deletion warnings. |
| Confirmation panel | `fraConfirm`, `lblConfirmIntro`, `lblConfirmSummary`, `chkConfirmAction`, `lblConfirmError` | Presents the exact action and requires confirmation for destructive work. |
| Result panel | `fraResult`, `lblResultSummary` | Displays the final result. |
| Navigation | `btnBack`, `btnCancel`, `btnOK` | Goes back, cancels, advances/executes, or closes according to the active panel. |

### 4.8 `frmSeedInventory` — Seed Demo Inventory

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblTitle` | Label — Seed demo inventory into which warehouse? | Explains the target choice. |
| `lblWarehouse`, `cmbWarehouse` | Label and combo box — Warehouse | Lists valid current target options and retains WarehouseId/station/root/status metadata. |
| `lblStation`, `txtStation` | Label and text box — Station | Supplies the station, defaulting to `S1`. |
| `lblUser`, `txtUser` | Label and text box — Admin user | Supplies the invSys user for the seed event. |
| `lblRoot`, `lblRootValue` | Labels — Runtime root | Shows the selected runtime root. |
| `lblStatus` | Status label | Shows target readiness, validation, or inbox-repair result. |
| `btnRepairInboxes` | Button — **Repair Inboxes** | Ensures Receiving, Shipping, and Production inboxes exist for the station. |
| `btnOK` | Button — **OK** | Accepts the selection; the public callback then queues and processes the seed event. |
| `btnCancel` | Button — **Cancel** | Cancels without seeding. |

Current acceptance status: the form and callback return, but the required three
entities are not visible after operator refresh. This form is therefore not
release-accepted.

### 4.9 `frmSetupTesterStation` — Setup Tester Station

This form reuses the Create Warehouse designer blob. It hides all designer
controls first, then selectively reuses fields and adds new prompt/actions.

Visible input mapping:

| Prompt/control | Reused or runtime input | Purpose |
|---|---|---|
| `lblUserIdPrompt` | `txtAdminUser` | Tester invSys UserId. |
| `lblPinPrompt` | `txtWarehouseName` | PIN input; the reused field is masked. |
| `lblConfirmPinPrompt` | `txtConfirmPinRuntime` | Confirms the PIN; masked. |
| `lblWarehousePrompt` | `txtWarehouseId` | Tester warehouse identifier. |
| `lblStationPrompt` | `txtStationId` | Tester station identifier. |
| `lblPathLocalPrompt` | `txtPathLocal` | NAS warehouse hub path. |
| `lblSharePointPrompt` | `txtPathSharePoint` | Locally synced SharePoint root. |

Other visible controls:

| Control | Displayed text | Purpose |
|---|---|---|
| `lblAdminUserError`, `lblWarehouseNameError`, `lblConfirmPinErrorRuntime`, `lblWarehouseIdError`, `lblStationIdError`, `lblPathLocalError`, `lblPathSharePointError` | Inline errors | Report validation errors for the mapped fields. |
| `lblSummary` | Status label | Shows setup/delete/open progress and final report. |
| `btnHubPathHelperRuntime` | **Find...** | Browses for the NAS hub. |
| `btnSharePointHelperRuntime` | **Find...** | Detects or browses for the SharePoint root. |
| `btnDeleteGeneratedRuntime` | **Delete Generated** | Deletes only guarded tester-generated artifacts. |
| `btnOpenReceivingRuntime` | **Open Workbook** | Appears after success and opens the generated Receiving workbook. |
| `btnOK` | **Setup** or **Close** | Runs setup, then becomes Close. |
| `btnCancel` | **Close** | Closes the form. |

Hidden but still present from the shared designer blob:
`lblIntro`, all six original `*Caption` labels, `chkPublishInitial`, and the
unused original layout positions. They are not part of the visible tester
workflow.

### 4.10 Admin shells with no deployed controls

| Form | Current source state |
|---|---|
| `frmAdminControls` | Empty designer in the deployed package; handlers reference absent `btnCreateDeleteUser` and `btnEditUser`. |
| `frmAdminEmail` | Empty designer and no executable form code. |
| `frmEditUser` | Empty designer; code references legacy controls that are not present in the deployed form. |
| `ufAdminItemSearch` | Empty role-named shell with no current caller. |
| `ufDynItemSearchTemplate` | Empty template shell; it has no controls unless a dynamic controller adds them. |

## 5. Core forms

### 5.1 `frmSignIn` — invSys Sign In

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblTitle` | Label — Sign in to invSys | Form heading. |
| `lblTarget` | Context label | Shows selected warehouse and station or that no server is connected. |
| `lblStatus` | Status label | Shows authentication guidance/errors. |
| `lblUser`, `txtUser` | Label and text box — User ID | Accepts the invSys user ID. |
| `lblSecret`, `txtSecret` | Label and masked text box — PIN/password | Accepts the invSys credential. |
| `btnOK` | Button — **Sign In** | Authenticates against the selected warehouse Auth workbook. |
| `btnCancel` | Button — **Cancel** | Cancels sign-in. |

### 5.2 `frmWarehouseConnection` — Connect Warehouse Storage

| Area | Controls | Purpose |
|---|---|---|
| Heading/status | `lblTitle`, `lblStatus` | Shows **Connect Warehouse Storage** and connection guidance/results. |
| Root | `lblRoot`, `txtRoot`, `btnScan` | Accepts and scans a NAS/server warehouse root. |
| Credentials | `lblUser`, `txtUser`, `lblPassword`, `txtPassword`, `btnConnect` | Connects the Windows session to the server; password is masked. |
| Station | `lblStation`, `cboStation`, `chkRequireStation`, `lblStationHelp` | Selects station and controls whether a station inbox is mandatory. |
| Targets | `lblTargets`, `lstTargets` | Lists discovered warehouse runtimes and their readiness. |
| Confirmation | `btnOK`, `btnCancel` | Selects the highlighted target or cancels. |

### 5.3 Dynamic item-search window (`cDynItemSearch`)

The controller adds the following four controls to a role-named empty form
shell such as `ufProductionItemSearch` or `ufShippingItemSearch`:

| Control | Type | Purpose |
|---|---|---|
| `txtSearch` | Text box | Filters the loaded role/mode-specific list. |
| `chkShippable` | Check box — **Shipping items only** or **Shippable only** | Restricts results for Shipping/component modes. |
| `lst` | Multi-column list box | Displays search results; selection/double-click/keyboard behavior routes the chosen item. |
| `txtDesc` | Read-only multiline text box | Displays description/details for the current result. |

The form caption varies by context: Receiving, Shipping, Production, Admin,
Production Ingredient, Production Palette Item, Production Recipe, or Shipping
Component Search.

### 5.4 Deprecated/incomplete Core form

`frmItemSearch` is explicitly deprecated in source. Its code expects `txtBox`
and `lstBox`, but its current deployed designer has zero controls. It must not
be treated as a working alternative to `cDynItemSearch`.

## 6. Receiving forms

### 6.1 `frmReceiving` — Receiving

| Area | Controls | Purpose |
|---|---|---|
| Tabs | `tabsReceiving` — **Receiving**, **Purchasing** | Switches between the working Receiving page and a reserved Purchasing page. |
| Receipt identity | `lblReceiptId`, `txtReceiptId` | Shows a generated, locked receipt ID. |
| Reference | `lblRef`, `txtRef` | Accepts PO/BOL reference text. |
| Inventory filter | `lblSearch`, `txtSearch` | Filters visible inventory as the user types. |
| Quantity | `lblQty`, `txtQty` | Accepts quantity for the selected item; defaults to 1. |
| Top actions | `btnRefresh`, `btnAdd` | Reloads views or stages the selected item/quantity. |
| Inventory | `lblInventoryTitle`, `lblInventoryHeader`, `lstInventory` | Displays System_Key, code, item, UOM, inventory, location, and description/vendor. |
| Staged receipt | `lblStagedTitle`, `lblStagedHeader`, `lstStaged` | Displays local Received Tally rows: reference, item, quantity, System_Key. |
| Aggregate view | `lblAggregateTitle`, `lblAggregateHeader`, `lstAggregate` | Displays aggregated reference/code/vendor/description/item/UOM/quantity/location/System_Key data. |
| Write actions | `btnConfirm`, `btnClear` | **Confirm Writes** queues supported Receiving events; **Clear** clears local staging. |
| Status/exit | `txtStatus`, `btnClose` | Shows multiline status and closes the form. |
| Purchasing placeholder | `lblPurchasingStub` | States that Purchasing is not operational and exposes no purchasing write action. |

### 6.2 Receiving shells

| Form | Current source state |
|---|---|
| `frmReceivingSavedList` | Empty packaged shell; no deployed controls. |
| `ufReceivingItemSearch` | Empty role shell; no current controller invocation in Receiving. |
| `ufDynItemSearchTemplate` | Duplicate source shell; not imported into the consolidated Operations project. |

## 7. Production forms

### 7.1 `frmProduction` — shared shell

| Control | Type | Purpose |
|---|---|---|
| `mpProduction` | MultiPage | Contains four pages: Recipe Builder, Ingredients Assignment, Production Run - List, and Production Run - Tree. |
| `txtProductionStatus` | Locked multiline text box | Shows bound workbook, inventory/design authority, validation, and action status. |
| `btnProductionClose` | Button — **Close** | Closes the Production form. |

### 7.2 Recipe Builder page

| Control group | Controls | Purpose |
|---|---|---|
| Saved recipes | `lstBuilderRecipes` | Lists recipe ID/name/description choices. |
| Recipe fields | `txtRecipeName`, `txtRecipeId`, `txtRecipeRowBudget`, `txtRecipeDescription` | Edit recipe identity, row budget, and description. |
| Recipe actions | `btnBuilderRefresh`, `btnBuilderNew`, `btnBuilderLoad`, `btnBuilderSave`, `btnBuilderProcess`, `btnBuilderFormulas`, `btnBuilderClear`, `btnBuilderRelease` | Refresh, start, load, save, add process tables, save formulas, clear, or release a recipe. |
| Line fields | `txtLineProcess`, `cmbLineIo`, `txtLineIngredient`, `txtLinePercent`, `cmbLineUom`, `txtLineAmount` | Edit process, input/output classification, ingredient/output/instruction, percent, UOM, and amount. |
| Line actions | `btnLineUomAdd`, `btnLineAdd`, `btnLineUpdate`, `btnLineRemove`, `btnLineMoveUp`, `btnLineMoveDown` | Add UOM, add/update/remove a recipe line, or reorder it. |
| Builder grid | `lstBuilderLines` | Lists Process, I/O, ingredient/output/instruction, %, UOM, amount, and ingredient ID. |
| Labels | Saved Recipes, Recipe Name, Recipe ID, Row Budget, Description, Process, In/Out, Ingredient / Output / Instruction, Percent, UOM, Amount, Recipe Builder Lines | Identify the controls and generated list headers. |

Generated list-header control families are `hdrBuilderLines1` through
`hdrBuilderLines8`, `hdrLoaderRecipes1` through `hdrLoaderRecipes3`,
`hdrLoaderLines1` through `hdrLoaderLines8`, `hdrRunPalette1` through
`hdrRunPalette10`, `hdrManagerCheck1` through `hdrManagerCheck6`, and
`hdrManagerOutput1` through `hdrManagerOutput8`. Blank or hidden data columns
still receive a header control where the runtime builder creates one.

### 7.3 Ingredients Assignment page

| Control group | Controls | Purpose |
|---|---|---|
| Recipes | `lstAssignRecipes`, `btnAssignRecipe`, `btnAssignRefresh` | Lists recipes, selects one, and reloads data. |
| Ingredients | `lstAssignIngredients`, `btnAssignIngredient` | Lists recipe ingredients and selects one for assignment. |
| Assignment actions | `btnAssignSave`, `btnAssignClear` | Saves or clears the current assignment context. |
| Inventory search | `txtInventorySearch`, `lstAssignInventory` | Filters and lists candidate inventory. |
| Allowed choices | `lstAssignAllowed`, `btnAssignAdd`, `btnAssignRemove` | Lists acceptable substitutions and adds/removes rows. |
| Labels | Recipes, Recipe Ingredients, Search Inventory, Inventory, Acceptable Items | Identify the page sections. |

### 7.4 Production Run - List page

| Control group | Controls | Purpose |
|---|---|---|
| Recipe loader | `lstLoaderRecipes`, `lstLoaderLines`, `btnLoaderRefresh`, `btnLoaderLoad`, `btnLoaderClear` | Selects a released recipe, shows its lines, refreshes, loads, or clears a run. |
| Run inputs | `cmbRunProcess`, `cmbRunLocation`, `txtPaletteSplit`, `txtPaletteQty`, `btnRunApplyPalette` | Chooses process/location and applies either percent-of-requirement or explicit quantity to the selected palette row. |
| Palette | `lstRunPalette` | Lists ingredient, System Key, inventory choice, requirement %, quantity, UOM, inventory, and location. |
| Inventory check | `lstManagerCheck` | Lists System Key, code, item, UOM, used quantity, and total inventory. |
| Output | `lstManagerOutput`, `txtOutputReal` | Lists process/output/UOM/last/batch/total/recall/inventory ID and accepts real output. |
| Run actions | `btnManagerCheckIn`, `btnManagerApplyOutput`, `btnManagerRefresh`, `btnManagerNext`, `btnManagerPrint` | Checks inputs in, completes the run, refreshes, advances to the next batch, or prints recall data. |
| Labels | Recipes, Loaded Recipe Lines, Process, Run Location, % of Requirement, Qty, Acceptable Inventory For Run, Inventory Check, Production Output, Real Output | Identify the page sections and generated headers. |

### 7.5 Production Run - Tree page

| Control group | Controls | Purpose |
|---|---|---|
| Tree choices | `lstRunTree` | Displays hierarchical recipe ingredient/inventory choices. |
| Tree inputs | `txtTreePaletteSplit`, `txtTreePaletteQty`, `cmbTreeRunProcess`, `cmbTreeRunLocation`, `btnRunTreeApplyPalette` | Applies percent/quantity/process/location to the selected tree choice. |
| Tree navigation | `btnRunTreeExpandAll`, `btnRunTreeCollapseAll` | Expands or collapses all tree groups. |
| Generated headers | Ingredient, System Key, Inventory Item, % Req, Qty, UOM, Inv, Location | Identify tree columns. |

`mBtnManagerPrepare`, `mBtnManagerUsed`, `mBtnManagerMade`, and
`mBtnManagerTotal` are declared event variables with handlers but are not
constructed by the current layout. They are not current visible controls.

Current visual acceptance status: **failed**. The native window can maximize,
but the screenshot shows the page/control layout retaining its base-size
footprint.

### 7.6 Production shells

`frmCreateRecipeTable`, `frmCreateSubstitutionList`, `frmIngredientPalette`,
and `frmSubstitution` are packaged empty shells with no controls.
`ufProductionItemSearch` receives the four dynamic search controls described in
section 5.3. `ufDynItemSearchTemplate` is a duplicate source shell and is not
imported into the consolidated Operations project.

## 8. Shipping forms

### 8.1 `frmShipmentsTally` — shared shell and tabs

| Control | Type | Purpose |
|---|---|---|
| `tabsShippingRole` | TabStrip — **Shipping**, **Box Builder**, **Box Maker** | Selects the active Shipping sub-workflow. |
| `txtStatus` | Locked multiline text box | Displays validation, queue, history, and action status. |
| `btnClose` | Button — **Close** | Closes the form. |

### 8.2 Shipping tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/history | `lblTitle`, `btnHistory`, `btnHistorySheet`, `btnRefresh` | Shows Shipments, opens history, exports history to a sheet, and refreshes data. |
| Search/filter | `lblPicker`, `txtPicker`, `chkUseExisting`, `lblSyncState` | Filters boxes, optionally uses existing shippable inventory, and shows pending/complete sync state. |
| Shippables | `lstShippables` plus `hdrShipBox`, `hdrShipVersion`, `hdrShipInv`, `hdrShipProjected`, `hdrShipLocked`, `hdrShipUom`, `hdrShipLoc`, `hdrShipRow` | Lists Box, Version, NAS Inv, Projected Inv, Locked, UOM, Location, and the current prohibited ROW display. |
| Line editor | `txtRef`, `txtBox`, `txtVersion`, `txtQty`, `txtUom`, `txtLocation`, `txtRow`, `txtCarrier`, hidden `txtDescription` | Edits reference, quantity, carrier, and selection-backed box/version/UOM/location/legacy ROW fields. |
| Line labels | `lblRef`, `lblBox`, `lblVersion`, `lblQty`, `lblUom`, `lblLocation`, `lblRow`, `lblCarrier` | Identify the editor fields. |
| Line actions | `btnAdd`, `btnUpdate`, `btnRemove` | Adds, updates, or removes a local shipment row. |
| Shipments list | `lblShipments`, `lstShipments` | Displays active shipment staging rows. |
| Shipment headers | generated `hdrRef*`, `hdrLineBox*`, `hdrLineQty*`, `hdrLineUom*`, `hdrLineArea*`, `hdrLineLocked*`, `hdrLineRow*`, `hdrLineDesc*`, `hdrLineCarrier*` | Identify Ref, Box, Qty, UOM, Area, Locked, legacy ROW, Version, and Carrier. One header set is built for each shipment/hold list. |
| Shipment actions | `btnStage`, `btnSend`, `btnHold` | Moves rows to shipment staging, sends completed shipments, or places selected rows on hold. |
| Hold list | `lblHold`, `lstHold`, `btnReturn` | Displays Not Shipped rows and returns selected rows to active staging. |

**D14 conflict:** `hdrShipRow`, `lblRow`, `txtRow`, and both generated
`hdrLineRow*` labels are architecturally invalid. Their current locked display
does not make `ROW` permissible.

The source also contains an uncalled readiness-header builder for
`hdrReadyType`, `hdrReadyItem`, `hdrReadyReq`, `hdrReadyInv`,
`hdrReadyStaged`, `hdrReadyUom`, `hdrReadyLoc`, `hdrReadyRow`, and
`hdrReadyStatus`. These are not constructed by the current form layout;
`hdrReadyRow` would be another D14 violation if that helper were activated.

### 8.3 Box Builder tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/actions | `lblBoxBuilderPage`, `btnBoxBuilderNewPage`, `btnBoxBuilderRefreshPage` | Starts a box definition and reloads box designs. |
| Designs | `lstBoxBuilderDesignsPage` | Lists design identity/name/version/status metadata. |
| Component inventory | `lblBoxBuilderInventory`, `lstBoxBuilderInventoryPage`, `lblBoxBuilderComponentQty`, `txtBoxBuilderComponentQty`, `btnBoxBuilderAddComponentPage`, `btnBoxBuilderRemoveComponentPage` | Selects inventory components and quantities for the box definition. |
| Definition fields | `lblBoxBuilderName`, `txtBoxBuilderName`, `lblBoxBuilderVersion`, `cboBoxBuilderVersion`, `lblBoxBuilderStatus`, `cboBoxBuilderStatus`, `lblBoxBuilderUom`, `txtBoxBuilderUom`, `lblBoxBuilderLocation`, `txtBoxBuilderLocation`, `lblBoxBuilderDescription`, `txtBoxBuilderDescription` | Edits box name, version, Active/Archived status, UOM, location, and description. |
| Selected components | `lblBoxBuilderComponents`, `lstBoxBuilderComponentsPage` | Displays the component set for the selected version. |
| Save/version actions | `btnBoxBuilderSavePage`, `btnBoxBuilderUpdateVersionPage`, `btnBoxBuilderNewVersionPage`, `btnBoxBuilderDeleteVersionPage`, `btnBoxBuilderArchivePage`, `btnBoxBuilderDeletePage` | Saves the box, updates or creates a version, deletes a version, archives the box, or deletes it through guarded paths. |

### 8.4 Box Maker tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/refresh | `lblBoxMakerPage`, `btnBoxMakerRefreshPage` | Identifies Box Maker and reloads released designs/inventory state. |
| Designs/version | `lstBoxMakerDesignsPage`, `lblBoxMakerVersion`, `cboBoxMakerVersion` | Selects a released box design and version. |
| Quantity/actions | `lblBoxMakerQty`, `txtBoxMakerQty`, `btnBoxMakerMakePage`, `btnBoxMakerUnmakePage` | Accepts quantity and makes boxes or unboxes them through the event path. |
| Components | `lblBoxMakerComponents`, `lstBoxMakerComponentsPage` | Displays selected-version component requirements and availability. |

### 8.5 `frmBoxVersionSaveChoice`

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblTitle` | Label — Save Box Version | Heading. |
| `lblBody` | Wrapped label | Explains the choice for the current box version. |
| `btnUpdate` | Button — **Update Version** | Selects in-place update behavior where allowed. |
| `btnNewVersion` | Button — **New Version** | Selects append-new-version behavior. |
| `btnCancel` | Button — **Cancel** | Cancels without saving. |

### 8.6 Shipping shells

`frmShippingCreateList` and `frmShippingSavedList` are packaged empty shells.
`ufShippingItemSearch` receives the four dynamic search controls described in
section 5.3. `ufDynItemSearchTemplate` is a duplicate source shell and is not
imported into the consolidated Operations project.

## 9. Review checklist

When reading this catalog, decide for each active form:

1. Is every control necessary for Release 1?
2. Is the displayed wording clear to an operator who does not know the VBA
   module names?
3. Should the control be visible to all signed-in users or only to a specific
   capability?
4. Does the control mutate only local staging, queue an event, or perform an
   administrative action, and is that distinction visible?
5. Does the control use `System_Key` whenever one durable inventory entity is
   identified?
6. Does resizing preserve readable proportions rather than merely enlarging
   the native form window?
7. Is an empty shell still needed, or should it be retired only after reviewed
   reachability and regression evidence?

## 10. Known required follow-up

- Create a focused packaged seed-round-trip RED that exercises
  `modAdmin.Seed_DemoInventory`, processor snapshot publication, and refresh of
  a captured operator workbook; require exactly three new visible entities.
- Create native-window Production geometry RED through
  `mProduction.BtnOpenProductionForm`, not only by assigning `Me.Width` and
  `Me.Height` inside the form.
- Replace Shipping `ROW` controls and any backing authority path with
  `System_Key` under D14 and focused packaged form-action tests.
- Do not delete empty/deprecated shells solely from this inventory. Apply the
  repository dead-code policy first.
