Attribute VB_Name = "importDataEntrySheet"
Public Sub importDataEntrySheet()
    Dim Filename As String
    Dim Sheet As Worksheet
    Dim wbImport As Workbook
    Dim dumpingTo As Workbook
    
    Application.ScreenUpdating = False
    Set dumpingTo = ActiveWorkbook

    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Please select a file to open:"
        .Filters.Add "Excel files", "*.xls; *.xls*", 1
        .Show
        On Error Resume Next 'In case the user has clicked the cancel button
            Filename = .SelectedItems(1)
            If Err.Number <> 0 Then
                Exit Sub 'Error has occurred so quit
            End If
        On Error GoTo 0
    End With
    
    Set wbImport = Workbooks.Open(Filename)
    dumpingTo.Sheets("Data Entry").Range("A1:F70") = wbImport.Sheets(1).Range("A1:F70").Value
    wbImport.Close

    Application.ScreenUpdating = True

End Sub
