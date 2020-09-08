Attribute VB_Name = "imageImport"
Sub imageImport()
'
' imageImport Macro
'
    Dim PicList() As Variant
    Dim PicFormat As String
    Dim Rng As Range
    Dim sShape As Shape
    Dim gridSize As Integer
    Dim headerText As String
    
    'User Input
    On Error Resume Next
    headerText = Application.InputBox(prompt:="Header Text: ", Type:=2)
    PicList = Application.GetOpenFilename(PicFormat, MultiSelect:=True)
    gridSize = Application.InputBox(prompt:="GridSize: (2-5)", Type:=1)
    
    
    If gridSize > 0 Then
        'Col Sizing
        Dim i As Integer: i = 1
        Dim width As Integer: width = 76 / gridSize
        Do While i <= gridSize
            Columns(i).ColumnWidth = width
            i = i + 1
        Loop
        
        'Row Sizing
        i = 1
        Dim rowNum As Integer: rowNum = 1
        Do While i * ((width * 7.2) + (width * 2)) <= 705 'PAGESIZE=705 MAX=409
            Rows(rowNum).RowHeight = width * 7.2
            Rows(rowNum + 1).RowHeight = width
            Rows(rowNum + 2).RowHeight = width
            i = i + 1
            rowNum = rowNum + 3
        Loop
    
        'Image Insertion
        xColIndex = 1
        If IsArray(PicList) Then
            xRowIndex = 1
            For lLoop = LBound(PicList) To UBound(PicList)
                Set Rng = Cells(xRowIndex, xColIndex)
                Set sShape = ActiveSheet.Shapes.AddPicture(PicList(lLoop), msoFalse, msoCTrue, Rng.Left, Rng.Top, Rng.width, Rng.height)
                If xColIndex = gridSize Then
                    xRowIndex = xRowIndex + 3
                    xColIndex = 1
                Else
                    xColIndex = xColIndex + 1
                End If
            Next
        End If
    End If
    
    ActiveSheet.PageSetup.CenterHeader = "&20" & headerText
    With Range("$A$1:$E$15")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 9 * (width / 15)
        .Font.Bold = True
    End With
    
    Range("$A$3:$E$3").Font.Size = 7 * (width / 15)
    Range("$A$3:$E$3").Font.Bold = False
    Range("$A$6:$E$6").Font.Size = 7 * (width / 15)
    Range("$A$6:$E$6").Font.Bold = False
    Range("$A$9:$E$9").Font.Size = 7 * (width / 15)
    Range("$A$9:$E$9").Font.Bold = False
    Range("$A$12:$E$12").Font.Size = 7 * (width / 15)
    Range("$A$12:$E$12").Font.Bold = False
    Range("$A$15:$E$15").Font.Size = 7 * (width / 15)
    Range("$A$15:$E$15").Font.Bold = False
    
End Sub

