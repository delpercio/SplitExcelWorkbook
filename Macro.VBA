Sub CreatingFilesFromSheets()
Dim folderPath As String
Dim sh As Worksheet
Dim fileExplorer As FileDialog
Dim wb As Workbook
Dim saveFileName As String
Set fileExplorer = Application.FileDialog(msoFileDialogFolderPicker)
' To allow or disable multi-select
fileExplorer.AllowMultiSelect = False
With fileExplorer
    If .Show = -1 Then ' A folder is selected
        folderPath = .SelectedItems.Item(1) ' Get the folder path
    Else ' The dialog is cancelled
        MsgBox "No folder selected. Operation cancelled.", vbInformation + vbOKOnly, "Cancelled"
        Exit Sub
    End If
End With
' Disable events
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each sh In ThisWorkbook.Sheets
    sh.Copy
    Set wb = ActiveWorkbook
    saveFileName = folderPath & "\" & sh.Name & ".xlsx"
    On Error Resume Next
    wb.SaveAs Filename:=saveFileName, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "Error saving sheet '" & sh.Name & "': " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    On Error GoTo 0
    wb.Close SaveChanges:=False
Next
' Enable events
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Sheets saved successfully!", vbInformation + vbOKOnly, "Success"
End Sub
