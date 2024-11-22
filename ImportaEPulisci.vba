Option Explicit


Private Sub Importa()
'
' Importa Macro
' Importa un file .txt
'
Dim mioFile As String
'
'    ChDir "C:\Users\gtorr\OneDrive\Desktop\Esercizi Excel\dataset"
mioFile = Application.GetOpenFilename("Text Files (*.txt), *.txt")
    Workbooks.OpenText Filename:= _
        mioFile, Origin _
        :=xlMSDOS, StartRow:=4, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0 _
        , 1), Array(8, 1), Array(20, 1), Array(26, 1), Array(41, 1), Array(49, 1), Array(59, 1), _
        Array(67, 1)), DecimalSeparator:=".", ThousandsSeparator:=",", _
        TrailingMinusNumbers:=True
    Rows("2:2").Delete Shift:=xlUp
    Range("A1").Select
End Sub
Private Sub Riempi()
'
' Riempi Macro
' Riempie le celle vuote con il valore sopra
'

'
    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
'    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C" 'Questo è un commento
    Range("A1").Select
End Sub
Private Sub IncollaValori()
'
' IncollaValori Macro
' Toglie le formule e lascia i valori
'

'
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
End Sub

Sub ImportaePulisci()
On Error Resume Next

Importa
Riempi
IncollaValori
Tabella
CreaPivot
Salva

End Sub
Sub Tabella()
'
' Tabella Macro
' Formatta come tabella
'

'
Dim tbl As ListObject

    Selection.CurrentRegion.Select
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.TableStyle = "TableStyleLight9"
    Range("Tabella1[#All]").Select
    
End Sub


Sub CreaPivot()
'
' CreaPivot Macro
' Crea una tabella pivot
'

'
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Tabella1", Version:=6).CreatePivotTable TableDestination:="Foglio1!R3C1", _
        TableName:="Tabella pivot2", DefaultVersion:=6
    Sheets("Foglio1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("Tabella pivot2")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
'        .PrintDrillIndicators = False
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("Tabella pivot2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("Tabella pivot2").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("Tabella pivot2").PivotFields("State")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabella pivot2").PivotFields("Channel")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabella pivot2").AddDataField ActiveSheet.PivotTables( _
        "Tabella pivot2").PivotFields("Gross"), "Somma di Gross", xlSum
    Range("B7").Select
    With ActiveSheet.PivotTables("Tabella pivot2").PivotFields("Somma di Gross")
        .NumberFormat = "#.##0,00"
    End With
End Sub
Sub Salva()
'
' Salva Macro
' Salva la cartella di lavoro corrente come file .xlsx
'
Dim fName As String
'
   fName = Application.GetSaveAsFilename
   ActiveWorkbook.SaveAs Filename:=fName


End Sub
