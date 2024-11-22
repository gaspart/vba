Sub ChangeCase()

' Cambia il testo in maiuscolo, minuscolo o maiuscola iniziale

Dim Rng As Range

    For Each Rng In Selection.Cells
        If Rng.HasFormula = False Then
            Rng.Value = LCase(Rng.Value) 'tutto minuscolo
        End If
    Next Rng

End Sub
