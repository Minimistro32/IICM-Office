Attribute VB_Name = "MergeWorkbooks"
Public Sub ImportCompRoster()
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
    
    ' Copy over Information into the MissionaryRoster Sheet
    Set wbImport = Workbooks.Open(Filename)
    wbImport.Sheets(1).UsedRange.Copy dumpingTo.Sheets("MissionaryRoster").Range("A1")
    wbImport.Close
    

    Call CleanCompRoster
    Application.ScreenUpdating = True

End Sub
Private Sub CleanCompRoster()
    Dim SourceRange As Range
    Dim EntireColumn As Range
    Dim EntireRow As Range
 
    On Error Resume Next
 
    Set SourceRange = Sheets("MissionaryRoster").UsedRange
 
    If Not (SourceRange Is Nothing) Then
        Application.ScreenUpdating = False
        ActiveSheet.Cells.UnMerge

        For i = SourceRange.Rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(i, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) < 4 Then
                EntireRow.Delete
            End If
        Next

        For i = SourceRange.Columns.Count To 1 Step -1
            Set EntireColumn = SourceRange.Cells(1, i).EntireColumn
            If Application.WorksheetFunction.CountA(EntireColumn) = 0 Then
                EntireColumn.Delete
            End If
        Next
        
        
        Application.ScreenUpdating = True
    End If
End Sub
