Sub PulisciTesto()
    Dim rng As Range
    Dim Area As Range
    Dim cell As Range

    ' Weed out any formulas from selection
    If Selection.Cells.Count = 1 Then
        Set rng = Selection
    Else
        Set rng = Selection.SpecialCells(xlCellTypeConstants)
    End If

    ' Trim and Clean cell values
    For Each Area In rng.Areas
        Area.Value = Evaluate("IF(ROW(" & Area.Address & "),CLEAN(TRIM(" & Area.Address & ")))")
    Next Area

    ' Iterate again to convert to ProperCase
    For Each Area In rng.Areas
        For Each cell In Area
            If Not IsNumeric(cell.Value) Then ' Ensure the cell contains text
                cell.Value = StrConv(cell.Value, vbProperCase)
            End If
        Next cell
    Next Area

End Sub
