Sub HalloWorld()

Dim Msg, Style, Title, Help, Ctxt, Response, MyString

Msg = "Hello World! You happy?"    ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2    ' Define buttons.
'Style = vbInformation + vbDefaultButton2    ' Define buttons.
Title = "Sciao belli!"    ' Define title.
Help = "DEMO.HLP"    ' Define Help file.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Good!"    ' Perform some action.
Else    ' User chose No.
    MyString = "Sorry about that"    ' Perform some action.
End If
MsgBox MyString
End Sub
