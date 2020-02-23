Sub SaveWorkshetAsPDF()

' Crea un PDF da ogni foglio della cartella corrente
Dim ws As Worksheet

    For Each ws In Worksheets
        ws.ExportAsFixedFormat xlTypePDF, "C:Users\Username\Desktop\Test\" & ws.Name & ".pdf"
    Next ws

End Sub
