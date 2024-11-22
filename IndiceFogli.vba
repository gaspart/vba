Sub IndiceFogli()

    'Crea un nuovo foglio all'inizio della cartella di lavoro corrente, lo rinomina
    'e vi inserisce in colonna A i link a ogni foglio di lavoro,
    'escludendo i fogli nascosti;
    'alla fine dell'elenco, aggiunge i nomi dei fogli grafico, senza link.
    
    'Attenzione:
    'i nomi dei fogli di lavoro non devono contenere spazi o caratteri speciali.
    
    Dim wSheet As Worksheet
    Dim oChart As Chart

    
    ActiveWorkbook.Sheets.Add(before:=Worksheets(1)).Name = "Indice" 'Un nome a caso
    Range("A1").Select
    
    For Each wSheet In Worksheets
        If wSheet.Visible = xlSheetVisible Then 'Controlla che il foglio sia visibile
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", _
            SubAddress:=wSheet.Name & "!A1", TextToDisplay:=wSheet.Name
            ActiveCell.Offset(1, 0).Select 'Scende di una riga
        End If
    Next wSheet
    
    For Each oChart In Application.Charts
        ActiveCell = oChart.Name
        ActiveCell.Offset(1, 0).Select 'Scende di una riga
    Next oChart
    
    Range("A1").Activate 'Riparte da A1
    Range(Selection, Selection.End(xlDown)).Font.Size = 18 'Seleziona da A1 alla fine e cambia la dimensione del font
    Range("A1").EntireColumn.AutoFit 'Adatta la colonna
    Range("A1").EntireRow.Delete 'Rimuove dall'elenco il foglio indice
    
End Sub
