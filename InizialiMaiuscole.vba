Sub ChangeCase()

' Cambia il testo in maiuscolo, minuscolo o maiuscola iniziale

Dim Rng As Range

    For Each Rng In Selection.Cells
        If Rng.HasFormula = False Then
            Rng.Value = Application.WorksheetFunction.Proper(Rng.Value) 'Iniziali Maiuscole
        End If
    Next Rng

End Sub
