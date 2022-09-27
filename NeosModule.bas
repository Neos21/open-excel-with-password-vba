Attribute VB_Name = "NeosModule"
' Open Excel File With Password
Sub OpenExcelFileWithPassword(filePath As String, filePass As String)
  On Error Resume Next
  Workbooks.Open Filename:=filePath, Password:=filePass
  If Err.Number <> 0 Then
    MsgBox "Failed To Open The Book"
  End If
End Sub
