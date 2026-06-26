Found it. The answer is clear and specific.

***

The clipboard pollution is coming from **`RebuildTableAtSurface`** in `modRoleWorkbookSurfaces.bas`, and it's being triggered on the **Arctic-Raptor workbook** (the operator's local `.xlsb`) whenever `EnsureShippingWorkbookSurface` or a similar surface-ensure call runs.

## What's Happening

`RebuildTableAtSurface` reads a table's `.Range.Value` into a `data` variant, calls `lo.Unlist`, clears the source range, then writes `targetRange.Value = data`. **That `.Clear` + `targetRange.Value = data` sequence does not dirty the clipboard.** However, `RebuildTableAtSurface` ends with:

```vba
CleanExit:
    On Error Resume Next
    Application.CutCopyMode = False  ' ← this clears march-ants but implies .Copy was called somewhere
```

The clipboard is actually being populated one step earlier in the **call chain**. When `ArrangeShippingBackendTablesSurface` fires (triggered by a surface-ensure call like `BtnOpenShipmentsForm`), it does:

```vba
MoveTableTopLeftSurface ws, "AggregateBoxBOM_Log", "CO1"
MoveTableTopLeftSurface ws, "AggregatePackages_Log", "CY1"
```

Both `AggregateBoxBOM_Log` and `AggregatePackages_Log` have headers **`GUID, USER, ACTION, ROW, ITEM_CODE, ITEM, ...`** — exactly what you see in your clipboard.

Similarly, `EnsureAllWorkbookSurfaces` → `EnsureProductionWorkbookSurface` calls:
```vba
EnsureTableSurface wb, "Production", "RecipeBuilder", Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", ...)
```

Which is your second clipboard block.

The real culprit is **`EnsureTableSurface` when the table doesn't exist yet** — it does this:

```vba
For i = LBound(headers) To UBound(headers)
    startCell.Offset(0, i - LBound(headers)).Value = headers(i)  ' ← writes cells directly
Next i
Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
```

**`ListObjects.Add` with `xlSrcRange` on an already-populated range leaves Excel's clipboard in a "copy mode" state on some Office builds.** Excel internally marks the source range as copied during table creation, and if `CutCopyMode = False` isn't called immediately after (or the `On Error Resume Next` swallows a failure), the header row stays in the clipboard.

## Why It's the Arctic-Raptor Machine Specifically

Arctic-Raptor is running the operator workbook that **doesn't yet have these log tables created** (or they got moved/deleted), so `EnsureTableSurface` is executing the table-creation branch — not the "table already exists, just ensure columns" branch. The machine where the tables already existed won't hit this path.

## Fix

In `EnsureTableSurface`, after the `ListObjects.Add` call, add an explicit clipboard clear:

```vba
Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
lo.Name = tableName
Application.CutCopyMode = False   ' ← add this line here
```

And in `RebuildTableAtSurface`, move the `CutCopyMode = False` **before** `CleanExit` so it runs even on the success path, not just error exit:

```vba
    newLo.ShowTotals = showTotals
    Application.CutCopyMode = False   ' ← move here, before CleanExit label
CleanExit:
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub
```

Once Arctic-Raptor's log tables exist (`AggregateBoxBOM_Log`, `AggregatePackages_Log`, `RecipeBuilder`, etc.), the creation branch won't fire again and the clipboard pollution will stop on its own — but the fix above prevents it on any fresh machine going forward.
			
			
