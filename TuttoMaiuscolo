Sub ChangeCase()

' Cambia il testo in maiuscolo, minuscolo o maiuscola iniziale

Dim Rng As Range

    For Each Rng In Selection.Cells
        If Rng.HasFormula = False Then
            Rng.Value = UCase(Rng.Value) 'TUTTO MAIUSCOLO
        End If
    Next Rng

End Sub
