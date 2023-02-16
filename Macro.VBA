Sub SplitWorkbook()
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim newWb As Workbook
    Dim filename As String
    Dim numErrors As Integer
    ' Unprotect all sheets and unhide all hidden sheets
    For Each sh In ThisWorkbook.Worksheets
        sh.Unprotect
        sh.Visible = xlSheetVisible ' unhide hidden sheets
    Next sh
    ' Disable screen updating and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wb = ActiveWorkbook
    numErrors = 0
    ' Loop through all sheets
    For Each sh In wb.Sheets
        ' Create a new workbook
        Set newWb = Workbooks.Add
        ' Copy the current sheet to the new workbook
        On Error Resume Next
        sh.Copy Before:=newWb.Sheets(1)
        If Err.Number <> 0 Then
            numErrors = numErrors + 1
        Else
            ' Delete all other sheets in the new workbook
            Application.DisplayAlerts = False
            While newWb.Sheets.Count > 1
                newWb.Sheets(2).Delete
            Wend
            Application.DisplayAlerts = True
            ' Set the filename for the new workbook
            filename = wb.Path & "\" & sh.Name & ".xlsx"
            ' Save and close the new workbook
            newWb.SaveAs filename, FileFormat:=xlOpenXMLWorkbook
            newWb.Close SaveChanges:=False
        End If
        On Error GoTo 0
    Next sh
    ' Re-enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If numErrors = 0 Then
        MsgBox "All sheets have been split into individual workbooks."
    Else
        MsgBox "There were " & numErrors & " sheets that could not be copied."
    End If
End Sub
