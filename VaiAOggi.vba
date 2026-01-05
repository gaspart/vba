Sub ScrollToTodayOrPreviousInCurrentColumn()
    Dim todayDate As Date
    Dim colRange As Range
    Dim cell As Range
    Dim lastMatch As Range

    todayDate = Date

    ' Define the range from the first used row to the last used row in the active column
    With ActiveCell.EntireColumn
        Set colRange = Range(.Cells(1), .Cells(.Worksheet.Rows.Count).End(xlUp))
    End With

    ' Loop through each cell and find the closest date <= today
    For Each cell In colRange
        If IsDate(cell.Value) Then
            If cell.Value <= todayDate Then
                Set lastMatch = cell
            Else
                Exit For ' Exit loop as dates are sorted and this one is greater than today
            End If
        End If
    Next cell

    If Not lastMatch Is Nothing Then
        Application.Goto lastMatch, True
    Else
        MsgBox "No date before or equal to today (" & Format(todayDate, "dd/mm/yyyy") & ") found in column.", vbInformation
    End If
End Sub
