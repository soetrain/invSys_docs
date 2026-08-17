# invSys Form Controls v1

**Version:** 1.8

**Inventory date:** 2026-08-17

**Architecture:** invSys v4.11, Release 1
**Scope:** every checked-in VBA UserForm source file used by Core, Admin,
Operations, Receiving, Production, or Shipping, including active controls,
runtime-created controls, status surfaces, and generated column labels.

## 1. Purpose and authority

This is the readable Release 1 control catalog after review-checklist answers
1-3 and 6-7. It describes the implemented source and explicitly labels packaged
or visible acceptance still pending. The normative authority remains
`invSys-Design-v4.11.md`.

The catalog was derived from:

- the 17 retained `.frm` files under `src` after reviewed reachability cleanup;
- all four checked-in `.frx` designer blobs;
- the current deployed five-package set in `deploy/current`;
- runtime control-construction calls and their event handlers; and
- the 2026-08-04 and 2026-08-08 operator screenshots and acceptance reports.

The retained forms produce 17 distinct components. Item search uses the single
Core-owned `frmItemSearch` runtime canvas plus role profiles in
`Core.cDynItemSearch`; no role-named or duplicate template shell remains.

This document covers forms only. Ribbon controls and worksheet buttons are
outside its scope. The older `controls.md` remains a broader, historical
status list; this file supersedes it only for the form-control inventory.

### Status terms

| Term | Meaning |
|---|---|
| Active | A current launcher or workflow constructs and uses the form. |
| Runtime-generated | The `.frm` has an empty designer and creates controls in VBA. |
| Designer-backed | Controls are stored in a checked-in `.frx` blob. |
| Stub | A visible, intentional Release 1 placeholder, currently Purchasing in Receiving. |

## 2. Acceptance findings that affect this catalog

### Role launcher/session stability: visible checkpoint passed

Against `WHT7025AE` / `S1`, repeated Receiving, Production, and Shipping
launches reused their respective saved workbooks and forms; no additional
operator workbook opened. Activating another workbook did not rebind the forms,
and minimizing the bound workbook minimized its form with it. After restarting
Excel and repeating the launches, the session remained stable. The verified
workbook names are `WHT7025AE.Receiving.Operator.xlsm`,
`WHT7025AE.Production.Operator.xlsm`, and
`WHT7025AE.Shipping.Operator.xlsm`.

### Seed Demo Inventory: packaged correction GREEN; dedicated UAT pending

`frmSeedInventory` displayed and its public callback reported one applied
processor event. The operator then refreshed, but the three demo entities were
not visible in that checkpoint. Subsequent read-only inspection found the demo
entities in all three saved role projections, and the packaged test proved the
same three unique `System_Key` values with `Condition=GOOD` through
canonical inventory, snapshot, saved Receiving workbook, and the Receiving
form's actual Refresh handler. That result is insufficient for a complete
Receiving-to-Production-to-Shipping workflow. The source correction now queues
one 24-entity batch containing raw ingredients, WIP, finished/shippable goods,
cases/boxes, tins, shipping cartons, dividers, labels, tape, and void fill,
with new `System_Key`, `Condition=GOOD`, and catalog
metadata. The packaged public callback is now GREEN through the event,
canonical inventory, catalog, published snapshot, saved Receiving projection,
and the Receiving form's actual Refresh handler for the complete 24-row kit.
All 24 entities were visible, all keys were unique, every condition was `GOOD`,
and the isolated packaged Receiving -> Production -> Box Maker -> Shipping ->
restart/reconciliation chain passed 30/30. The 2026-08-09 dedicated NAS
checkpoint confirmed that the visible Seed Inventory action added inventory;
the expanded 24-row kit still requires the visible dedicated-NAS checkpoint.

### Production resizing: corrected; visible retest pending

`frmProduction` maximized its native window while the MultiPage and child
controls remained near their base dimensions in the upper-left corner. The
focused validator reproduced the failure at DPI-derived `Zoom=60`. The deployed
correction holds the form at `Zoom=100`; packaged validation now passes native
minimize/restore/maximize/restore and bounds/fill/overlap checks for all four
pages. A visible maximize/restore retest remains required.

### Shipping status/message resizing: packaged correction GREEN; visible retest pending

The 2026-08-08 screenshots show `txtStatus` initially above Search Boxes and
then translated below Search Boxes after a height resize. It was incorrectly
anchored Left + Right + Bottom. It is now anchored Left + Top + Right, so its
Top and Height remain fixed while its Width follows the form. The focused
source contract and the packaged public-launcher grow/shrink proof are GREEN.
A visible grow/shrink/maximize/restore retest remains pending.

### Inventory Viewer: Release 1 stub packaged GREEN

Operations now contains a signed-in, capability-neutral **Inventory Viewer**
entry. It opens one reusable modeless form and reads only the current published
inventory snapshot on explicit refresh. It does not open a Receiving,
Production, or Shipping form and never writes, repairs, processes, or refreshes
an authority workbook. Packaged validation proved repeated launch reused the
same form generation, local filtering reduced the list to the matching row,
three snapshot levels loaded, and the inspected snapshot hash did not change.
The Operations ribbon entry now uses a visible built-in Excel table icon.

### Shipping visible identity path: `System_Key` correction packaged GREEN

`frmShipmentsTally` now exposes `System Key` in the shippable list, line editor,
shipment list, and hold list. The form passes the exact string key to the
captured-workbook staging, reservation, hold/return, and Shipments Sent paths;
the event creator emits `System_Key` without a managed `ROW` field. The focused
source contract is 12/12 GREEN, and the packaged public Shipping launcher proof
confirms that the control preserves a nonnumeric key and that the reservation
key is derived from it. Private legacy worksheet-maintenance routines are not
represented as Release 1 controls or authority paths; they remain subject to
separate reachability review before removal.

## 3. Complete form inventory

| Package/source area | Form | Current status | Control surface |
|---|---|---|---|
| Admin | `frmAddInventoryItem` | Active, runtime-generated | Add/edit inventory metadata and custom fields |
| Admin | `frmAdminDesignLifecycle` | Active, runtime-generated | Design refresh/import/release/obsolete actions |
| Admin | `frmAdminSettings` | Active, runtime-generated | Config, connection policy, carriers, UOM catalog |
| Admin | `frmCreateDeleteUser` | Active, runtime-generated | Warehouse user and capability maintenance |
| Admin | `frmCreateWarehouse` | Active, designer-backed plus runtime buttons | Greenfield warehouse creation |
| Admin | `frmReAuthGate` | Active, designer-backed | Destructive-action re-authentication |
| Admin | `frmRetireMigrateWarehouse` | Active, designer-backed plus runtime controls | Archive, retire, migrate, or delete warehouse |
| Admin | `frmSeedInventory` | Active, runtime-generated | Choose target and seed demo inventory |
| Admin | `frmSetupTesterStation` | Active, reused designer plus runtime controls | Disposable tester-station setup and cleanup |
| Core | `frmItemSearch` | Active runtime canvas | Controls and role profiles supplied by `cDynItemSearch` |
| Core | `frmSignIn` | Active, runtime-generated | invSys user authentication |
| Core | `frmWarehouseConnection` | Active, runtime-generated | NAS/root connection and warehouse target selection |
| Operations | `frmInventoryViewer` | Active, runtime-generated | Read-only current inventory levels, search, freshness, refresh |
| Production | `frmProduction` | Active, runtime-generated | Recipe, assignment, run-list, and run-tree workflows |
| Receiving | `frmReceiving` | Active, runtime-generated | Receiving and non-operational Purchasing tab |
| Shipping | `frmBoxVersionSaveChoice` | Active, runtime-generated | Choose update-versus-new box alternative behavior |
| Shipping | `frmShipmentsTally` | Active, runtime-generated | Shipping, Box Designer, and Box Maker tabs |

## 4. Admin forms

### 4.1 `frmAddInventoryItem` — Add/Edit Inventory Item

This form edits inventory item metadata and queues the supported add/edit
workflow. It does not make a worksheet row the durable identity.

| Control | Type / displayed text | Purpose |
|---|---|---|
| `btnAddMode` | Button — **Add Item Mode** | Selects add mode and clears add fields. |
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

The Admin ribbon exposes one **Design Lifecycle** launcher. Release and
Obsolete are actions inside this form, not separate ribbon launchers.

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
| `txtStationId` | Text box — Station ID | Creates the initial station, defaulting to the Windows computer name. |
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
| `lblStation`, `txtStation` | Label and locked text box — Station | Shows the Windows computer name used automatically as station identity. |
| `lblUser`, `txtUser` | Label and text box — Admin user | Supplies the invSys user for the seed event. |
| `lblRoot`, `lblRootValue` | Labels — Runtime root | Shows the selected runtime root. |
| `lblStatus` | Status label | Shows target readiness, validation, or inbox-repair result. |
| `btnRepairInboxes` | Button — **Repair Inboxes** | Ensures Receiving, Shipping, and Production inboxes exist for the station. |
| `btnOK` | Button — **OK** | Accepts the selection; the public callback then queues and processes the seed event. |
| `btnCancel` | Button — **Cancel** | Cancels without seeding. |

Current acceptance status: the prior three-row seed was not visible in the
user's Receiving refresh and was not a complete workflow kit. The rebuilt
package now requests and publishes 24 entities, including five explicit
box-making consumables. The public Admin callback and isolated packaged
Receiving/Production/Box Maker/Shipping full chain are GREEN; visible
dedicated-NAS acceptance remains required.

### 4.9 `frmSetupTesterStation` — Test Environment Setup

This is an Admin-only isolated test-environment provisioner, not the normal
station identity mechanism and not merely a diagnostics runner. It is retained
because it creates and cleans up disposable warehouse/operator fixtures used by
packaged regression tests. The form reuses the Create Warehouse designer blob. It hides all designer
controls first, then selectively reuses fields and adds new prompt/actions.

Visible input mapping:

| Prompt/control | Reused or runtime input | Purpose |
|---|---|---|
| `lblUserIdPrompt` | `txtAdminUser` | Tester invSys UserId. |
| `lblPinPrompt` | `txtWarehouseName` | PIN input; the reused field is masked. |
| `lblConfirmPinPrompt` | `txtConfirmPinRuntime` | Confirms the PIN; masked. |
| `lblWarehousePrompt` | `txtWarehouseId` | Tester warehouse identifier. |
| `lblStationPrompt` | `txtStationId` | Test station identifier; ordinary workflows use the Windows computer name. |
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

### 4.10 Reviewed Admin cleanup

Reachability showed that `frmAdminControls`, `frmAdminEmail`, `frmEditUser`,
`ufAdminItemSearch`, and `ufDynItemSearchTemplate` had no current public
launcher or useful deployed control surface. They were removed before the
package rebuild. The nine active Admin forms above are the Release 1 set.

## 5. Core forms

### 5.1 `frmSignIn` — invSys Sign In

The form authenticates only against the current committed warehouse target.
Changing warehouse, station, or runtime root signs out the prior session so a
credential can never be silently checked against the previously selected
warehouse. The target line must show the selected Warehouse ID and this
computer's Windows name before the user enters the warehouse-specific
credential.

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
| Station | `lblStation`, locked `txtStation`, `lblStationHelp` | Shows the Windows computer name. The role contract still decides whether a station inbox is required; the user does not choose a bespoke station. |
| Targets | `lblTargets`, `lstTargets` | Lists discovered warehouse runtimes and their readiness. |
| Confirmation | `btnOK`, `btnCancel` | Selects the highlighted target or cancels. |

The Operations **Send To** selector and this form share the same target-binding
contract. If the selected warehouse still contains only a legacy station such
as `S1`, selecting it may add this computer's Windows name to station config
after the warehouse Auth workbook validates. Arbitrary typed or stale station
names are still rejected. The selected target is committed only after all
validation succeeds; a failure reports the attempted target and explicitly
states that the prior current target remains unchanged.

Automated 2026-08-17 evidence is GREEN for the exact ribbon action path,
current-computer enrollment, cross-target sign-out, authentication binding, and
two clean dedicated-NAS sessions against `WHT7025AE` / `X1-PRO-AI`. The visible
operator sign-in with the user's existing warehouse credential remains the next
acceptance action.

### 5.3 Dynamic item-search window (`cDynItemSearch`)

The controller adds the following four controls to the clean Core-owned
`frmItemSearch` runtime canvas. Callers select a role profile rather than a
role-named form:

| Control | Type | Purpose |
|---|---|---|
| `txtSearch` | Text box | Filters the loaded role/mode-specific list. |
| `chkShippable` | Check box — **Shipping items only** or **Shippable only** | Restricts results for Shipping/component modes. |
| `lst` | Multi-column list box | Displays search results; selection/double-click/keyboard behavior routes the chosen item. |
| `txtDesc` | Read-only multiline text box | Displays description/details for the current result. |

The form caption varies by context: Receiving, Shipping, Production, Admin,
Production Ingredient, Production Palette Item, Production Recipe, or Shipping
Component Search.

### 5.4 Shared runtime canvas

`frmItemSearch` intentionally contains no designer controls or legacy event
code. `cDynItemSearch` owns all controls, search behavior, result routing, and
role-profile wording. This is the only Release 1 item-search form component.

## 6. Operations forms

### 6.1 `frmInventoryViewer` — Current inventory levels

Ribbon entry: Operations > Overview > **Inventory Viewer**. It is visible
without a role capability restriction, while its action requires a signed-in
invSys session and selected warehouse.

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblTitle` | Label — **Current inventory levels** | Identifies the read-only overview. |
| `btnRefresh` | Button — **Refresh** | Reads the current published inventory snapshot; it does not process or alter the snapshot. |
| `lblSearch`, `txtSearch` | Label and text box — **Search** | Filters the already loaded rows locally across all visible columns. |
| `lblHeaders` | Header label | Identifies Item Code, Item, UOM, Quantity, Location, and Condition. |
| `lstInventory` | Six-column list box | Displays inventory levels aggregated by item code, item, UOM, location, and condition. |
| `lblStatus` | Status/freshness label | Shows row count, snapshot read time, or a no-snapshot/sign-in error. |
| `btnClose` | Button — **Close** | Closes the Viewer without affecting an operator workbook. |

The modeless form is reused on repeated launch. It is resizable: Search,
headers, list, and status expand or reposition through the shared Operations
anchor manager while remaining readable.

## 7. Receiving forms

### 7.1 `frmReceiving` — Receiving

| Area | Controls | Purpose |
|---|---|---|
| Tabs | `tabsReceiving` — **Receiving**, **Purchasing** | Switches between the working Receiving page and a reserved Purchasing page. |
| Receipt identity | `lblReceiptId`, `txtReceiptId` | Shows a generated, locked receipt ID. |
| Reference | `lblRef`, `txtRef` | Accepts PO/BOL reference text. |
| History filter | `lblSearch`, `txtSearch` | Filters Receiving Entries History as the user types. |
| Managed item search | `lblItemSearch`, `txtItemSearch`, `lblReceiveItemsTitle`, `lblReceiveItemsHeader`, `lstReceiveItems` | Filters and displays a dedicated result list with Code, Item, UOM, Available, Location, Description, and Vendor while retaining hidden source `System_Key`. |
| Quantity | `lblQty`, `txtQty` | Accepts quantity for the selected item; defaults to 1. |
| Receipt attributes | `lblReceiveLocation`, `txtReceiveLocation`, `lblLotNumber`, `txtLotNumber` | Requires the receiving location and accepts an optional lot number. Selection defaults Location from the source row, but the operator may change it. |
| Top actions | `btnRefresh`, `btnAdd` | Reloads views or stages the selected item/quantity. |
| Receiving history | `lblInventoryTitle`, `lblInventoryHeader`, `lstInventory` | Displays completed `ReceivedLog` entries: date, user, reference, item, quantity, UOM, vendor, location, code, and lot; one hidden column retains `System_Key`. `EventId` remains in the workbook log. |
| Staged receipt | `lblStagedTitle`, `lblStagedHeader`, `lstStaged` | Displays local Received Tally rows: reference, item, quantity, location, lot, and hidden System_Key. |
| Aggregate view | `lblAggregateTitle`, `lblAggregateHeader`, `lstAggregate` | Displays aggregated reference/code/vendor/description/item/UOM/quantity/location/lot data with hidden System_Key. |
| Write actions | `btnConfirm`, `btnClear` | **Confirm Writes** queues supported Receiving events; **Clear** clears local staging. |
| Status/exit | `txtStatus`, `btnClose` | Shows multiline status and closes the form. |
| Purchasing placeholder | `lblPurchasingStub` | States that Purchasing is not operational and exposes no purchasing write action. |

### 7.2 Reviewed Receiving cleanup

`frmReceivingSavedList`, `ufReceivingItemSearch`, and
`ufDynItemSearchTemplate` were unreachable empty shells and are removed. The
intentional Purchasing stub remains inside the active `frmReceiving` because
the reviewed Release 1 workflow reserves Purchasing there.

## 8. Production forms

### 8.1 `frmProduction` — shared role form

| Control | Type | Purpose |
|---|---|---|
| `mpProduction` | MultiPage | Contains four pages: Recipe Builder, Ingredients Assignment, Production Run - List, and Production Run - Tree. |
| `txtProductionStatus` | Locked multiline text box | Shows bound workbook, inventory/design authority, validation, and action status. |
| `btnProductionClose` | Button — **Close** | Closes the Production form. |

### 8.2 Recipe Builder page

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

### 8.3 Ingredients Assignment page

| Control group | Controls | Purpose |
|---|---|---|
| Recipes | `lstAssignRecipes`, `btnAssignRecipe`, `btnAssignRefresh` | Lists recipes, selects one, and reloads data. |
| Ingredients | `lstAssignIngredients`, `btnAssignIngredient` | Lists recipe ingredients and selects one for assignment. |
| Assignment actions | `btnAssignSave`, `btnAssignClear` | Saves or clears the current assignment context. |
| Inventory search | `txtInventorySearch`, `lstAssignInventory` | Filters and lists candidate inventory. |
| Allowed choices | `lstAssignAllowed`, `btnAssignAdd`, `btnAssignRemove` | Lists acceptable substitutions and adds/removes rows. |
| Labels | Recipes, Recipe Ingredients, Search Inventory, Inventory, Acceptable Items | Identify the page sections. |

### 8.4 Production Run - List page

| Control group | Controls | Purpose |
|---|---|---|
| Recipe loader | `lstLoaderRecipes`, `lstLoaderLines`, `btnLoaderRefresh`, `btnLoaderLoad`, `btnLoaderClear` | Selects a released recipe, shows its lines, refreshes, loads, or clears a run. |
| Batch scaling | `txtBatchScalePercent`, `btnApplyBatchScale` | Applies a List-run batch scale from `0.001%` through `1000%`; `100%` preserves the released recipe quantities. |
| Run inputs | `cmbRunProcess`, `cmbRunLocation`, `txtPaletteSplit`, `txtPaletteQty`, `btnRunApplyPalette` | Chooses process/location and applies either percent-of-requirement or explicit quantity to the selected palette row. |
| Palette | `lstRunPalette` | Lists ingredient, System Key, inventory choice, requirement %, quantity, UOM, inventory, and location. |
| Inventory check | `lstManagerCheck` | Lists System Key, code, item, UOM, used quantity, and total inventory. |
| Output | `lstManagerOutput`, `txtOutputReal` | Lists process/output/UOM/last/batch/total/recall/inventory ID and accepts real output. |
| Run actions | `btnManagerCheckIn`, `btnManagerApplyOutput`, `btnManagerRefresh`, `btnManagerNext`, `btnManagerPrint` | Checks inputs in, completes the run, refreshes, advances to the next batch, or prints recall data. |
| Labels | Recipes, Loaded Recipe Lines, Process, Run Location, % of Requirement, Qty, Acceptable Inventory For Run, Inventory Check, Production Output, Real Output | Identify the page sections and generated headers. |

### 8.5 Production Run - Tree page

This page remains an experimental alternative and is not the Release 1
workflow focus. Slice 4j intentionally makes no Tree behavior change; Production
Run - List is the path to validate.

| Control group | Controls | Purpose |
|---|---|---|
| Tree choices | `lstRunTree` | Displays hierarchical recipe ingredient/inventory choices. |
| Tree inputs | `txtTreePaletteSplit`, `txtTreePaletteQty`, `cmbTreeRunProcess`, `cmbTreeRunLocation`, `btnRunTreeApplyPalette` | Applies percent/quantity/process/location to the selected tree choice. |
| Tree navigation | `btnRunTreeExpandAll`, `btnRunTreeCollapseAll` | Expands or collapses all tree groups. |
| Generated headers | Ingredient, System Key, Inventory Item, % Req, Qty, UOM, Inv, Location | Identify tree columns. |

The unconstructed `mBtnManagerPrepare`, `mBtnManagerUsed`,
`mBtnManagerMade`, and `mBtnManagerTotal` event variables/handlers were removed;
they were never visible controls. Current packaged geometry validation is GREEN
at `Zoom=100`; the user's visible maximize/restore retest remains pending.

### 8.6 Reviewed Production cleanup

`frmCreateRecipeTable`, `frmCreateSubstitutionList`, `frmIngredientPalette`,
`frmSubstitution`, `ufProductionItemSearch`, and
`ufDynItemSearchTemplate` were empty or unreachable shells and are removed.

## 9. Shipping forms

### 9.1 `frmShipmentsTally` — Shipping/Boxing role form and tabs

| Control | Type | Purpose |
|---|---|---|
| `tabsShippingRole` | TabStrip — **Shipping**, **Box Designer**, **Box Maker** | Selects the active Shipping sub-workflow. |
| `txtStatus` | Locked multiline text box | Displays validation, queue, history, and action status; remains fixed above Search Boxes during resize. |
| `btnClose` | Button — **Close** | Closes the form. |

### 9.2 Shipping tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/history | `lblTitle`, `btnHistory`, `btnHistorySheet`, `btnRefresh` | Shows Shipments, opens history, exports history to a sheet, and refreshes data. |
| Search/filter | `lblPicker`, `txtPicker`, `chkUseExisting`, `lblSyncState` | Filters boxes, optionally uses existing shippable inventory, and shows pending/complete sync state. |
| Shippables | `lstShippables` plus `hdrShipBox`, `hdrShipVersion`, `hdrShipInv`, `hdrShipProjected`, `hdrShipLocked`, `hdrShipUom`, `hdrShipLoc`, `hdrShipSystemKey` | Lists Box, Alternative, NAS Inv, Projected Inv, Locked, UOM, Location, and immutable System Key. |
| Line editor | `txtRef`, `txtBox`, `txtVersion`, `txtQty`, `txtUom`, `txtLocation`, `txtSystemKey`, `txtCarrier`, hidden `txtDescription` | Edits reference, quantity, and carrier while preserving the selection-backed box/alternative/UOM/location/System_Key identity. |
| Line labels | `lblRef`, `lblBox`, `lblVersion`, `lblQty`, `lblUom`, `lblLocation`, `lblSystemKey`, `lblCarrier` | Identify the editor fields. |
| Line actions | `btnAdd`, `btnUpdate`, `btnRemove` | Adds, updates, or removes a local shipment row. |
| Shipments list | `lblShipments`, `lstShipments` | Displays active shipment staging rows. |
| Shipment headers | generated `hdrRef*`, `hdrLineBox*`, `hdrLineQty*`, `hdrLineUom*`, `hdrLineArea*`, `hdrLineLocked*`, `hdrLineSystemKey*`, `hdrLineDesc*`, `hdrLineCarrier*` | Identify Ref, Box, Qty, UOM, Area, Locked, System Key, Alternative, and Carrier. One header set is built for each shipment/hold list. |
| Shipment actions | `btnStage`, `btnSend`, `btnHold` | Moves rows to shipment staging, sends completed shipments, or places selected rows on hold. |
| Hold list | `lblHold`, `lstHold`, `btnReturn` | Displays Not Shipped rows and returns selected rows to active staging. |

**D14 correction:** the visible form and its reachable Release 1 backing path
use exact string `System_Key`; no numeric value was relabeled as a key.

The uncalled readiness list and header builder were removed after reachability
review; they were never constructed by the active layout.

### 9.3 Box Designer tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/actions | `lblBoxBuilderPage`, `btnBoxBuilderNewPage`, `btnBoxBuilderRefreshPage` | Starts a box definition and reloads box designs. |
| Designs | `lstBoxBuilderDesignsPage` | Full-width list of design identity/name/alternative/status metadata. |
| Component inventory | `lblBoxBuilderInventory`, `lstBoxBuilderInventoryPage`, `lblBoxBuilderComponentQty`, `txtBoxBuilderComponentQty`, `btnBoxBuilderAddComponentPage`, `btnBoxBuilderRemoveComponentPage` | Selects inventory entities and quantities for the box definition. Each row retains its hidden immutable `System_Key`; repeated item names can represent distinct physical inventory entities rather than duplicate GUI data. |
| Component search | `txtBoxBuilderSearch` | Filters loaded component choices locally by code, item, description, UOM, or location. |
| Definition fields | `lblBoxBuilderName`, `txtBoxBuilderName`, `lblBoxBuilderVersion`, `cboBoxBuilderVersion`, `lblBoxBuilderStatus`, `cboBoxBuilderStatus`, `lblBoxBuilderUom`, `txtBoxBuilderUom`, `lblBoxBuilderLocation`, `txtBoxBuilderLocation`, `lblBoxBuilderDescription`, `txtBoxBuilderDescription` | Edits box name, alternative, Active/Archived status, UOM, location, and description. |
| Selected components | `lblBoxBuilderComponents`, `lstBoxBuilderComponentsPage` | Full-width list of the component set for the selected alternative. The hidden component identity is preserved as `System_Key` and persisted as `ComponentSystemKey` in the Shipping BOM. |
| Save/alternative actions | `btnBoxBuilderSavePage`, `btnBoxBuilderUpdateVersionPage`, `btnBoxBuilderNewVersionPage`, `btnBoxBuilderDeleteVersionPage`, `btnBoxBuilderArchivePage`, `btnBoxBuilderDeletePage` | Saves the box, updates or creates an alternative, deletes an alternative, archives the box, or deletes it through guarded paths. |

### 9.4 Box Maker tab

| Area | Controls | Purpose |
|---|---|---|
| Heading/refresh | `lblBoxMakerPage`, `btnBoxMakerRefreshPage` | Identifies Box Maker and reloads released designs/inventory state. |
| Designs/alternative | `lstBoxMakerDesignsPage`, `lblBoxMakerVersion`, `cboBoxMakerVersion` | Selects a released box design and alternative. |
| Quantity/actions | `lblBoxMakerQty`, `txtBoxMakerQty`, `btnBoxMakerMakePage`, `btnBoxMakerUnmakePage` | Accepts quantity and makes boxes or unboxes them through the event path. |
| Components | `lblBoxMakerComponents`, `lstBoxMakerComponentsPage` | Displays selected-alternative component requirements and availability resolved by the exact preserved `ComponentSystemKey`. Make/Unmake event payloads use those string keys and the package `System_Key`, never a numeric row surrogate. |

All Designer/Maker list boxes now use an explicit vertical, full-width layout
and participate in native form resizing without overlap. Each list has
a monospaced header strip whose position and width are recalculated from the
list geometry after anchor application. A managed item without a Shipping BOM
alternative displays `NA`; only a Shipping BOM/package alternative displays
the compatible stored labels `v1`, `v2`, and so on.

### 9.5 `frmBoxVersionSaveChoice`

| Control | Type / displayed text | Purpose |
|---|---|---|
| `lblTitle` | Label — Save Box Alternative | Heading. |
| `lblBody` | Wrapped label | Explains the choice for the current box alternative. |
| `btnUpdate` | Button — **Update Alternative** | Selects in-place update behavior where allowed. |
| `btnNewVersion` | Button — **New Alternative** | Creates another box design choice; this is an alternative, not a software release sequence. |
| `btnCancel` | Button — **Cancel** | Cancels without saving. |

### 9.6 Reviewed Shipping cleanup

`frmShippingCreateList`, `frmShippingSavedList`, `ufShippingItemSearch`, and
`ufDynItemSearchTemplate` were empty or unreachable shells and are removed.

## 10. Review checklist decisions

| Item | 2026-08-08 decision / state |
|---|---|
| 1. Necessary for Release 1? | Operations/Admin and the active role-form tabs are sufficient. Unreachable shells and unconstructed controls were removed. Inventory Viewer was added as the one missing overview surface; Purchasing remains the intentional Receiving stub. |
| 2. Operator wording? | Accepted as good enough for Release 1. |
| 3. Visibility? | Controls may remain visible to all signed-in users for Release 1. Existing capability checks on mutating actions remain defense-in-depth; Inventory Viewer has no role-capability restriction. |
| 4. Mutation/queue/admin distinction? | **Release 1 decision:** Search, filters, history, Viewer Refresh, and role Refresh are read-only projections. Add/Update/Remove/Hold/Return and recipe/BOM editing change only captured-workbook staging until an explicit save/post action. Confirm Writes, Production completion, Make/Unmake, and Shipments Sent queue inventory events; the processor is the only inventory authority writer. Admin controls perform named administrative actions and retain capability/re-auth/audit gates. Status text should state `staged`, `queued`, `applied`, or `refreshed` rather than imply that a local edit immediately changed inventory. |
| 5. Durable identity uses `System_Key`? | **Accepted rule:** every physical inventory entity, inventory event/log row, reservation, shipment, Production allocation/output, and Shipping BOM package/component reference uses immutable `System_Key`. Aggregate quantity views may group by SKU/location and need not impersonate one key. Designs use their specialized three-digit base-36 `DesignId` plus `DesignVersion`; an inventory entity produced from a design still receives its own `System_Key`. `EventId`, shipment IDs, and run IDs identify their own records. The reachable Shipping form/event path now follows this rule and preserves string identity. |
| 6. Readable resize proportions? | Required. Production, Shipping `txtStatus`, Viewer, and the explicit full-width Box Designer/Box Maker layouts are source- and packaged-geometry GREEN. Visible maximize/restore/grow/shrink confirmation remains part of the dedicated UAT checkpoint. |
| 7. Empty shells? | Retired after source reachability review and focused regression contracts. Only the visible Purchasing stub remains intentionally. |

## 11. Required follow-up to close the control review

- Repeat visible Production maximize/restore and Shipping/Boxing grow/shrink checks.
- Run the complete 24-row seed-to-ship workflow on the dedicated test warehouse,
  including Receiving Location/Lot and Production List scaling.
- Review and adjust checklist item 4 only if an operator test exposes an
  ambiguous status or mutation boundary.
- Review private legacy Shipping worksheet-maintenance routines for reachability
  before removal; they are outside the packaged Release 1 form/action authority
  path and must not be reintroduced as compatibility behavior.
