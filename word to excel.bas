Attribute VB_Name = "Module1"
Sub importtabledatawork()
Dim WdApp As Object, wddoc As Object
Dim strDocName As String
On Error Resume Next
Set WdApp = GetObject(, “Word.Application”)
If Err.Number = 429 Then
Err.Clear
Set WdApp = CreateObject(“Word.Application”)
End If
WdApp.Visible = True
strDocName = "C:\Users\ytdon\Desktop\lala.docx"
If Dir(strDocName) = "" Then
MsgBox "The file" & strDocName & vbCrLf & _
"was not found in the folder path" & vbCrLf & _
"C:\our-inventory\.", _
vbExclamation, _
"Sorry, that document name does not exist."
Exit Sub
End If

WdApp.Activate
Set wddoc = GetObject(strDocName)
If wddoc Is Nothing Then Set wddoc = WdApp.Documents.Open(strDocName)
wddoc.Activate

Dim Tble As Integer
Dim rowWd As Long
Dim colWd As Integer
Dim x As Long, y As Long
x = 1
y = 1

With wddoc
Tble = wddoc.Tables.Count
If Tble = 0 Then
MsgBox "No Tables found in the Word document", vbExclamation, "No Tables to Import"
Exit Sub
End If

For i = 1 To Tble
With .Tables(i)
    For rowWd = 1 To .Rows.Count
    For colWd = 1 To .Columns.Count
    Cells(x, y) = WorksheetFunction.Clean(.cell(rowWd, colWd).Range.Text)
    y = y + 1
    Next colWd
    y = 1
    x = x + 1
    Next rowWd
    End With
    Next
End With

wddoc.Close Savechanges:=False
WdApp.Quit
Set wddoc = Nothing
Set WdApp = Nothing

End Sub



