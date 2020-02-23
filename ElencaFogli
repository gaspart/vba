Sub ListSheets()
 
Dim ws As Worksheet
Dim x As Integer
 
x = 1
 
Sheets("Sheet1").Range("A:A").Clear
 
For Each ws In Worksheets
     Sheets("Sheet1").Cells(x, 1) = ws.Name
     x = x + 1
Next 'ws
 
End Sub
