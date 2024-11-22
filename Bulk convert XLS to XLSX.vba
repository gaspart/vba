Sub Convert_XLS_to_XLSX()
    Dim sourceFolder As String
    Dim destinationFolder As String
    Dim fileName As String
    Dim wb As Workbook
    Dim prohibitedChars As String
    Dim isValid As Boolean
    Dim i As Integer

    ' Define the source folder path (where the .xls files are located)
    sourceFolder = "C:\Users\gtorr\OneDrive\Documents\ESSEDI SRL\Pivot Vecchie"
    ' Define the destination folder path (where .xlsx files will be saved)
    destinationFolder = "C:\Users\gtorr\OneDrive\Documents\ESSEDI SRL\Pivot Nuove"

    ' Ensure the folder paths end with a backslash
    If Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"
    If Right(destinationFolder, 1) <> "\" Then destinationFolder = destinationFolder & "\"

    ' Define prohibited characters
    prohibitedChars = ":/\*?""<>|[]"

    ' Get the first .xls file in the source folder
    fileName = Dir(sourceFolder & "*.xls")

    ' Loop through all .xls files in the source folder
    Do While fileName <> ""
        isValid = True
        
        ' Check if the file name contains any prohibited characters
        For i = 1 To Len(prohibitedChars)
            If InStr(fileName, Mid(prohibitedChars, i, 1)) > 0 Then
                isValid = False
                Exit For
            End If
        Next i

        If isValid Then
            ' Open the .xls file
            Set wb = Workbooks.Open(sourceFolder & fileName)
            
            ' Save the workbook in .xlsx format in the destination folder, overwriting if file exists
            Application.DisplayAlerts = False
            wb.SaveAs fileName:=destinationFolder & Replace(fileName, ".xls", ".xlsx"), FileFormat:=xlOpenXMLWorkbook
            Application.DisplayAlerts = True
            
            ' Close the workbook without prompting to save changes
            wb.Close SaveChanges:=False
        Else
            ' Notify that a file was skipped
            Debug.Print "Skipped file due to prohibited characters: " & fileName
        End If

        ' Get the next .xls file
        fileName = Dir
    Loop

    MsgBox "Conversion completed!", vbInformation
End Sub
