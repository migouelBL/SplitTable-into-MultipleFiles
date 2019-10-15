Option Compare Database
Option Explicit

Private Sub Export_Click()
Dim Selected As Object
Dim FilePath, FileName, FileRec, Prefix, Response As String
Dim LastRecord, CurrentRecord, i, x As Integer

Application.Echo False
Me.FilterOn = False
On Error GoTo ErrorHandler

'-----------------Set Destination---------------
Set Selected = Application.FileDialog(4)
With Selected
    .AllowMultiSelect = False
    .Title = "Select Folder Destination"
    .Show
    i = .SelectedItems.Count
    If i = 0 Then GoTo Exiting
End With

'------------------Set Filenames-----------------
FilePath = Selected.SelectedItems(i) & "\"
FileName = Mid([Forms]![1_Main]![Path].Caption, InStrRev([Forms]![1_Main]![Path].Caption, "\") + 1)
FileName = Left(FileName, (InStrRev(FileName, ".", -1, vbTextCompare) - 1))
If [Forms]![2_SplitTable]![OGFileName] = False Then FileName = ""

Prefix = [Forms]![2_SplitTable]![Prefix]
Application.FollowHyperlink (FilePath)

'----------------Prepare Recordset---------------
CurrentRecord = 1
LastRecord = DCount("*", "Query1")
DoCmd.GoToRecord , , acLast

'-------------------Loop Export-----------------
For x = CurrentRecord To LastRecord
If CurrentRecord = LastRecord Then Exit For
    FileRec = Replace(Replace(Replace([Forms]![2_SplitTable]![Delimiter], "\", "-"), "/", "-"), ".", "_")
    DoCmd.TransferSpreadsheet transfertype:=acExport, _
    SpreadSheetType:=acSpreadsheetTypeExcel12Xml, _
    TableName:="Query2", _
    FileName:=FilePath & Prefix & FileRec & "_" & FileName & ".xlsx", _
    HasFieldNames:=True, Range:=[Forms]![2_SplitTable]![Delimiter]

DoCmd.GoToRecord , , acPrevious
CurrentRecord = CurrentRecord + 1
Next x

'-----------------Release Memory---------------
Exiting:
Set Selected = Nothing
FilePath = vbNullString
FileName = vbNullString
Application.Echo True

Exit Sub

'------------------Error Handler-----------------
ErrorHandler:
Application.Echo True
If Err.Number = 3191 Then
    MsgBox "Cannot proceed" & vbNewLine & _
    "Multiple field headers with the same name was found in the source file" & vbNewLine & _
    [Forms]![1_Main]![Path].Caption
ElseIf Err.Number = 3436 Then
    MsgBox "Cannot continue" & vbNewLine & _
    "Null values was found within: " & [Forms]![1_Main]![Delimiter]
Else
    MsgBox "Error No: " & Err.Number _
    & vbNewLine & "Description: " & Err.Description
End If
End Sub
