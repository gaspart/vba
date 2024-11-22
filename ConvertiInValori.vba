Sub ConvertToValues()

' Copia/incolla valori

    With ActiveSheet.UsedRange
        .Value = .Value
    End With

End Sub
