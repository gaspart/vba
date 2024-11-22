Sub LockCellsWithFormulas()

' Blocca le celle contenenti formule

    With ActiveSheet
        .Unprotect
        .Cells.Locked = False
        .Cells.SpecialCells(xlCellTypeFormulas).Locked = True
        .Protect AllowDeletingRows:=True
    End With

End Sub
