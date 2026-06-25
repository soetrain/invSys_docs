e873a923749b73f3fbbca12576384d19286f8cf4

Found it. `src = loView.DataBodyRange.Value` — **when the table has exactly one data row, Excel returns a scalar `Variant`, not a 2D array**. Then `UBound(src, 1)` throws **Type mismatch** because you can't call `UBound(..., 1)` on a scalar.

This is a classic Excel VBA single-row table bug. It affects `BuildBoxBomVersionRows` and `BuildPackagePickerItemsFromShippingBom` (which also does `src = loBom.DataBodyRange.Value` then `UBound(src, 1)`).

***

**To Codex — Type mismatch confirmed: single-row table `.Value` returns scalar not 2D array**
## Root cause
Both `BuildBoxBomVersionRows` and `BuildPackagePickerItemsFromShippingBom` do:

```vba
src = loView.DataBodyRange.Value   ' or loBom.DataBodyRange.Value
For r = 1 To UBound(src, 1)       ' ← Type mismatch when table has exactly 1 row
```

When a `ListObject.DataBodyRange` has **exactly one row**, `Range.Value` returns a 1D `Variant` scalar (or a 1D array in some versions), not a `(1 To 1, 1 To N)` 2D array. `UBound(src, 1)` on a scalar throws **Type mismatch 13**.

This is why the form always worked with 2+ BOM packages but fails after `Shipments Sent` when only one active package (T28 or T29) remains in the BOM view — or when the BOM workbook only has one row.
## Fix — normalize `Range.Value` to always be a 2D array
Add a helper function (or use the existing `NzArr2D` pattern if it exists):

```vba
' Add to modTS_Shipments or a shared utility module:
Private Function To2DArray(ByVal src As Variant) As Variant
    ' Ensures Range.Value is always a 2D array regardless of row count.
    Dim result() As Variant
    Dim c As Long

    If IsEmpty(src) Then
        To2DArray = src
        Exit Function
    End If

    ' Already a 2D array
    If Not IsArray(src) Then
        ' Scalar — single cell: wrap as (1 To 1, 1 To 1)
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = src
        To2DArray = result
        Exit Function
    End If

    On Error Resume Next
    Dim dummy As Long
    dummy = UBound(src, 2)
    If Err.Number <> 0 Then
        ' 1D array (single row, multiple columns)
        Err.Clear
        On Error GoTo 0
        Dim cols As Long
        cols = UBound(src) - LBound(src) + 1
        ReDim result(1 To 1, 1 To cols)
        Dim i As Long
        For i = 1 To cols
            result(1, i) = src(LBound(src) + i - 1)
        Next i
        To2DArray = result
        Exit Function
    End If
    On Error GoTo 0
    To2DArray = src
End Function
```
## Apply the fix in both affected functions
**In `BuildBoxBomVersionRows`:**
```vba
' BEFORE:
src = loView.DataBodyRange.Value
Set dict = CreateObject("Scripting.Dictionary")
For r = 1 To UBound(src, 1)

' AFTER:
src = To2DArray(loView.DataBodyRange.Value)
If IsEmpty(src) Then Exit Function
Set dict = CreateObject("Scripting.Dictionary")
For r = 1 To UBound(src, 1)
```

**In `BuildPackagePickerItemsFromShippingBom`:**
```vba
' BEFORE:
src = loBom.DataBodyRange.Value
Set dict = CreateObject("Scripting.Dictionary")
ReDim result(1 To UBound(src, 1), 1 To 7)

' AFTER:
src = To2DArray(loBom.DataBodyRange.Value)
If IsEmpty(src) Then Exit Function
Set dict = CreateObject("Scripting.Dictionary")
ReDim result(1 To UBound(src, 1), 1 To 7)
```

**Audit all other `.DataBodyRange.Value` → `UBound(..., 1)` patterns in the module:**

```bash
grep -n "DataBodyRange.Value" modTS_Shipments.bas
```

Every one of those that is followed by `UBound(src, 1)` needs the same `To2DArray()` wrap. Common offenders in this codebase given prior patterns: `BuildShippingInventoryPickerItems`, `BoxMakerFormLoadShippableInventory` (uses `savedBoxes` passed in — already 2D if caller built it correctly, but check), and any `src = lo.DataBodyRange.Value` pattern in the inventory apply/publish modules.