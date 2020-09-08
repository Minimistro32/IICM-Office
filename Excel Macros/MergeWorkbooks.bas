Attribute VB_Name = "MergeWorkbooks"
Function GFL(rng As Range) As String
    Dim arr
    Dim I As Long
    arr = VBA.Split(rng, " ")
    If IsArray(arr) Then
        For I = LBound(arr) To UBound(arr)
            GFL = GFL & Left(arr(I), 1)
        Next I
    Else
        GFL = Left(arr, 1)
    End If
End Function
Public Sub importOrgRoster()
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
    wbImport.Sheets(1).UsedRange.Copy dumpingTo.Sheets("Sheet1").Range("A1")
    wbImport.Close
    

    Call cleanOrgRoster
    Application.ScreenUpdating = True

End Sub
Private Sub cleanOrgRoster()
    Dim SourceRange As Range
    Dim EntireColumn As Range
    Dim EntireRow As Range
 
    On Error Resume Next
 
    Set SourceRange = Sheets("Sheet1").UsedRange
 
    If Not (SourceRange Is Nothing) Then
        Application.ScreenUpdating = False
        ActiveSheet.Cells.UnMerge

        For I = SourceRange.Rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(I, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) < 4 Then
                EntireRow.Delete
            End If
        Next

        'Delete trash cols and individual col processing
        For I = SourceRange.Columns.Count To 1 Step -1
            Set EntireColumn = SourceRange.Cells(1, I).EntireColumn
            If SourceRange.Cells(1, I).Value = "Last Name" Then
            ElseIf SourceRange.Cells(1, I).Value = "Status" Then
                For j = SourceRange.Rows.Count To 1 Step -1
                    If SourceRange.Cells(j, I).Value = "In Other Mission" Or SourceRange.Cells(j, I).Value = "Released" Then
                        SourceRange.Cells(j, I).EntireRow.Delete
                    End If
                Next
                EntireColumn.Delete
            ElseIf SourceRange.Cells(1, I).Value = "Position" Then
                ''For j = SourceRange.Rows.Count To 1 Step -1
                ''    If SourceRange.Cells(j, i).Value = "JC" Or SourceRange.Cells(j, i).Value = "SC" Then 'DL/SA/TR
                ''        SourceRange.Cells(j, i).Value = ""
                ''    End If
                ''Next
            ElseIf SourceRange.Cells(1, I).Value = "Phone1" Then
            ElseIf SourceRange.Cells(1, I).Value = "Phone2" Then
            ElseIf SourceRange.Cells(1, I).Value = "Phone3" Then
            ElseIf SourceRange.Cells(1, I).Value = "Zone" Then
                For j = SourceRange.Rows.Count To 1 Step -1
                    SourceRange.Cells(j, I).Value = Replace(SourceRange.Cells(j, I).Value, " Zone", "")
                Next
            ElseIf SourceRange.Cells(1, I).Value = "Area" Then
                For j = SourceRange.Rows.Count To 1 Step -1
                    If SourceRange.Cells(j, I).Value = "" Then
                        SourceRange.Cells(j, I).EntireRow.Delete
                    End If
                Next
            ElseIf SourceRange.Cells(1, I).Value = "Area Email" Then
            ElseIf SourceRange.Cells(1, I).Value = "Street" Then
            ElseIf SourceRange.Cells(1, I).Value = "City" Then
            ElseIf SourceRange.Cells(1, I).Value = "State/Province" Then
            ElseIf SourceRange.Cells(1, I).Value = "Postal Code" Then
            ElseIf SourceRange.Cells(1, I).Value = "Country" Then
            Else
                EntireColumn.Delete
            End If
        Next
        
        'get Col's new permenant positions and changing the names to Google's correct name
        Dim nameColNum
        Dim positionColNum
        Dim areaColNum
        Dim zoneColNum
        For I = SourceRange.Columns.Count To 1 Step -1
            Set EntireColumn = SourceRange.Cells(1, I).EntireColumn
            If SourceRange.Cells(1, I).Value = "Last Name" Then
                nameColNum = I
                SourceRange.Cells(1, I).Value = "Notes"
            ElseIf SourceRange.Cells(1, I).Value = "Position" Then
                positionColNum = I
                SourceRange.Cells(1, I).Value = "Name Prefix"
            ElseIf SourceRange.Cells(1, I).Value = "Area" Then
                areaColNum = I
                SourceRange.Cells(1, I).Value = "Given Name"
            ElseIf SourceRange.Cells(1, I).Value = "Area Email" Then
                SourceRange.Cells(1, I).Value = "E-mail 1 - Value"
            ElseIf SourceRange.Cells(1, I).Value = "Phone1" Then
                SourceRange.Cells(1, I).Value = "Phone 1 - Value"
            ElseIf SourceRange.Cells(1, I).Value = "Phone2" Then
                SourceRange.Cells(1, I).Value = "Phone 2 - Value"
            ElseIf SourceRange.Cells(1, I).Value = "Phone3" Then
                SourceRange.Cells(1, I).Value = "Phone 3 - Value"
            ElseIf SourceRange.Cells(1, I).Value = "Area" Then
                SourceRange.Cells(1, I).Value = "Given Name"
            ElseIf SourceRange.Cells(1, I).Value = "Zone" Then
                zoneColNum = I
                SourceRange.Cells(1, I).Value = "Group Membership"
            ElseIf SourceRange.Cells(1, I).Value = "Street" Then
                SourceRange.Cells(1, I).Value = "Address 1 - Street"
            ElseIf SourceRange.Cells(1, I).Value = "City" Then
                SourceRange.Cells(1, I).Value = "Address 1 - City"
            ElseIf SourceRange.Cells(1, I).Value = "State/Province" Then
                SourceRange.Cells(1, I).Value = "Address 1 - Region"
            ElseIf SourceRange.Cells(1, I).Value = "Postal Code" Then
                SourceRange.Cells(1, I).Value = "Address 1 - Postal Code"
            ElseIf SourceRange.Cells(1, I).Value = "Country" Then
                SourceRange.Cells(1, I).Value = "Address 1 - Country"
            End If
        Next
        
        'appending pos to name
        For j = 2 To SourceRange.Rows.Count Step 1
            If SourceRange.Cells(j, positionColNum).Value <> "" Then
                SourceRange.Cells(j, nameColNum).Value = SourceRange.Cells(j, positionColNum).Value & " " & SourceRange.Cells(j, nameColNum).Value
            End If
        Next
        
        'condense rows to unique areas
        For j = 1 To SourceRange.Rows.Count Step 1
            If SourceRange.Cells(j, areaColNum).Value <> "" Then
                For k = j + 1 To SourceRange.Rows.Count Step 1
                    If SourceRange.Cells(j, areaColNum).Value = SourceRange.Cells(k, areaColNum).Value Then
                        SourceRange.Cells(j, nameColNum).Value = SourceRange.Cells(j, nameColNum).Value & "; " & SourceRange.Cells(k, nameColNum).Value
                        SourceRange.Cells(k, areaColNum).EntireRow.Delete
                    End If
                Next
            End If
        Next
        
        
        For j = 2 To SourceRange.Rows.Count Step 1
            'change pos to name prefix
            Dim tempPrefixStorage As String
            tempPrefixStorage = GFL(SourceRange.Cells(j, zoneColNum))
            
            'update Group Membership
            SourceRange.Cells(j, zoneColNum).Value = "* myContacts ::: " & "@IICM ::: " & SourceRange.Cells(j, zoneColNum).Value
            If InStr(1, SourceRange.Cells(j, positionColNum).Value, "AP", vbTextCompare) Or InStr(1, SourceRange.Cells(j, positionColNum).Value, "ZL", vbTextCompare) Or InStr(1, SourceRange.Cells(j, positionColNum).Value, "STL", vbTextCompare) Then
                SourceRange.Cells(j, zoneColNum).Value = SourceRange.Cells(j, zoneColNum).Value & " ::: #MLC"
            End If
            
            SourceRange.Cells(j, positionColNum).Value = tempPrefixStorage
        Next
        
        Application.ScreenUpdating = True
        
        ActiveWorkbook.SaveAs Filename:= _
            "C:\Users\2011328\Desktop\Google Contact Import.csv", FileFormat:= _
            xlCSV, CreateBackup:=False
    End If
End Sub
