Sub SortSheetsTabName()

' Ordina i fogli in base al nome

Application.ScreenUpdating = False

Dim ShCount As Integer, i As Integer, j As Integer
ShCount = Sheets.Count

    For i = 1 To ShCount - 1
        For j = i + 1 To ShCount
            If Sheets(j).Name < Sheets(i).Name Then
                Sheets(j).Move before:=Sheets(i)
            End If
        Next j
    Next i

Application.ScreenUpdating = True

End Sub
