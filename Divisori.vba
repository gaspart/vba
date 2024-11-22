Sub Divisori()
Dim n As Double
Dim i As Double



n = 2
i = ActiveCell.Value

If Int(i) <> i Then
    MsgBox ("Solo numeri positivi interi")
    Exit Sub
End If

While i > 2 And n <= i
    If i Mod n <> 0 Then
        n = n + 1
    Else
        ActiveCell.Offset(1, 0).Select
        Selection = n
        i = i / n
    End If
Wend

End Sub
