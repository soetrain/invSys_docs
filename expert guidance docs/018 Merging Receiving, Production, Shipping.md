No new fetches needed — the full v4.10 text is already in hand from the last retrieval. Here is every place in the spec that assumes Receiving, Shipping, and Production are separate XLAMs, with the specific edit each one needs to reflect the `invSys.Operations.xlam` consolidation.

## New section required: an Architecture Decision entry

The spec has no decision record for this change at all, so before touching diagrams the plan needs a **new `D12 — Operations Packaging Consolidation`** entry, placed after D11 (Shipping Inventory Boundary) and before "System Topology." It should state:

- Receiving, Shipping, and Production combine into one deployed add-in, `invSys.Operations.xlam`, exposing one Excel ribbon tab ("Operations") with Receiving/Production/Shipping groups.
- Admin remains `invSys.Admin.xlam` as a separate ribbon tab.
- Core, `invSys.Inventory.Domain.xlam`, and `invSys.Designs.Domain.xlam` remain separate, headless (no ribbon tab).
- This is a **packaging-only** change: D3's Core/Domain/Role boundary is unchanged, D4's per-role form ownership is unchanged, and Receiving/Production/Shipping remain separate internal modules/forms inside the single XLAM.
- Rationale: reduces operator-visible Excel tabs from four to two and shortens the build/staging cycle, without merging source responsibilities.

Every other listed change below should cross-reference this decision so nothing looks like an unexplained contradiction to a future session.

## Sections needing direct edits

**Technology Stack — Distribution.** Currently reads "Deployment: XLAM add-ins + workbooks" with no count implied, so no change needed here beyond consistency, but the Release Strategy bullet "Role UIs: Receiving, Shipping, Production (VBA + RibbonX)" should get a parenthetical noting these ship inside one Operations XLAM.

**Repository Structure diagram.** The mermaid block currently shows `SRC --> RECV / SHIP / PROD` as three independent top-level folders, each implying its own XLAM boundary, with `RECV --> RECVF[Forms]` and `RECV --> RECVR[Ribbon]` shown per-role. This needs updating so the diagram shows a `SRC --> OPS[Operations]` node with `Receiving`, `Production`, `Shipping` as sub-folders/modules underneath it, plus a single shared `Ribbon` node at the Operations level (since there's now one ribbon tab, not three). This is the clearest place to visually encode "separate modules, one package."

**SharePoint Folder Structure diagram.** `ADDINSCURRENT` currently lists `XLAMRECV[invSys.Receiving.xlam]`, `XLAMSHIP[invSys.Shipping.xlam]`, and `XLAMPROD[invSys.Production.xlam]` as three separate nodes. These three need to collapse into a single `XLAMOPS[invSys.Operations.xlam]` node. This directly changes what gets published and versioned in `/invSys/Addins/Current`.

**Component Dependency Graph.** `RecvUI`, `ShipUI`, `ProdUI` are three separate styled nodes, each with its own edge from `Auth`. These should become sub-nodes or a single `OpsUI` node (with Receiving/Shipping/Production as internal groupings if the diagram needs that granularity), still fed by `Auth` and feeding `Proc` — the dependency relationships don't change, only the packaging boundary the nodes represent.

**D3 — Clear Ownership Boundaries.** The rule text is already boundary-agnostic ("Role XLAMs: UI + event creation only"), so no wording changes are strictly required, but it would help to add one clarifying sentence: "Receiving, Production, and Shipping are packaged together in `invSys.Operations.xlam` per D12, but each retains an independent UI/event-creation boundary as if it were still a separate XLAM." This preempts a future session assuming the consolidation loosened the boundary.

**D4 — Forms Strategy, Form Ownership Matrix.** The matrix's column headers (`Receiving | Shipping | Production | Admin`) currently read as if they map 1:1 to XLAM boundaries. They still work as-is for form ownership, but the row for `Core.ItemSearch` and the surrounding prose should note explicitly that "each role XLAM" in this section now means "each role module packaged inside `invSys.Operations.xlam`" — otherwise the phrase "each role XLAM contains: its own item-search userform" reads as contradicting D12.

**D-NAS — Rationale and Operator sign-in workflow.** The line "Receiving, Shipping, and Production operators may not have `invSys.Admin.xlam` loaded, so Admin cannot be the only place where NAS access is established" still holds true and needs no change. But item 6 in the sign-in workflow ("Role write/send buttons require an allowed warehouse target...") and the general framing of "Role XLAMs" throughout D-NAS should get the same one-line clarification as D3: these are Operations-XLAM-internal modules now, referenced collectively as "Role XLAMs" for boundary purposes.

**Phase 3 tasks.** "Build RibbonX XML for all role XLAMs" and "Build Receiving.UI + EventCreator / Build Shipping.UI + EventCreator / Build Production.UI + EventCreator" are historical, already-completed (`[x]`) tasks describing the pre-consolidation state. These should not be rewritten retroactively (they're a valid historical record), but a footnote should be added: "Superseded packaging: as of D12, these three UIs ship inside `invSys.Operations.xlam` rather than as separate XLAMs." This avoids a contributor reading Phase 3 and rebuilding three separate add-ins.

**Phase 6 tasks — the most operationally important edit.** The unchecked task "Move NAS connection handling, remembered warehouse target selection, and runtime resolver priority into Core... expose separate storage connection, invSys sign-in, sign-out, and current-user status controls from Receiving, Shipping, Production, and Admin ribbons" and its paired test both name four separate ribbons. Since Receiving/Production/Shipping now share one ribbon, this should be rewritten as: "...expose separate storage connection, invSys sign-in, sign-out, and current-user status controls from the Operations ribbon (shared across Receiving, Production, and Shipping groups) and the Admin ribbon." This matters because it's still an open Phase 6 gate — getting it wrong here would mean Codex builds and tests against a four-ribbon assumption that no longer matches the packaging plan.

**Phase 6 LAN operationalization / acceptance criteria.** The phrase "LAN station bootstrap... config, inbox, and shared-auth provisioning for the station user" and "Role-ready acceptance criteria" don't name XLAMs directly, so no change needed, but the "operator workflow dependability requirements" subsections ("Receiving is dependable on LAN only when...", "Shipping is dependable...", "Production is dependable...") should each get a shared preface note that these three now load as one XLAM, so a station only needs one add-in load/version check instead of three when validating the acceptance ladder.

**Tools section.** "`export-vba.ps1`, `build-xlam.ps1`" is currently unstated on build granularity. This is where the earlier "selective build mode" recommendation should be spec'd concretely: `build-xlam.ps1` needs a documented mode that builds only `invSys.Operations.xlam` (or only the changed module within it) plus any changed Core/Domain dependency, with the full five-XLAM package (Core, Inventory.Domain, Designs.Domain, Operations, Admin) built only at integration checkpoints. This is a new spec requirement, not just a diagram fix, and belongs either under a new subsection near "Repository Structure" or folded into D12.

## Summary of the packaging count change

| Spec artifact | Current state | Needed state |
|---|---|---|
| XLAM count | 7 (`Core`, `Inventory.Domain`, `Designs.Domain`, `Receiving`, `Shipping`, `Production`, `Admin`) | 5 (`Core`, `Inventory.Domain`, `Designs.Domain`, `Operations`, `Admin`) |
| Ribbon tabs | 4 visible (Receiving, Shipping, Production, Admin) | 2 visible (Operations, Admin) |
| Diagrams needing edits | Repository Structure, SharePoint Folder Structure, Component Dependency Graph | Same three, each with the three role nodes merged into one `Operations` node |
| New decision record | None | `D12 — Operations Packaging Consolidation` |
| Phase 6 tasks/tests referencing 4 ribbons | 1 task + 1 test | Rewritten to reference Operations + Admin (2 ribbons) |

Once D12 is drafted and these diagram/task edits are made, the spec will be internally consistent with the Operations-consolidation direction, and the exception no longer needs to be tracked informally — it becomes the normative document again, which is the position you and Codex both wanted it to hold.