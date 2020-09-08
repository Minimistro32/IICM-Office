Attribute VB_Name = "printCompHistory"
Function updPrintedRows() As Integer 'ByRef currPrintAreaRowNum As Integer)
    'Declaration
    Dim firstCol As Range, cell As Range, currPrintAreaRowNum As Integer
    
    'Init
    Set firstCol = Evaluate("$A$1:$A$" & Cells(Rows.Count, 1).End(xlUp).Row)
    currPrintAreaRowNum = 0
    
    'Loop through first column looking for "Total Records:"
    For Each cell In firstCol
        If Left(cell, 14) = "Total Records:" Then
            currPrintAreaRowNum = cell.Row
            Exit For
        End If
    Next cell
    
    'Return
    updPrintedRows = currPrintAreaRowNum
End Function

Sub printCompHistory()
Attribute printCompHistory.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Printcomphistory Macro
'

    'If there are people to print, process them
    Do While updPrintedRows() <> 0
        'Print
        Evaluate("$A$1:$K$" & currPrintAreaRowNum).PrintOut 'Debug.Print "Print A1:K" & updPrintedRows() & " - " & Range("A10")
        
        
        'Delete what was printed
        Rows("10:" & updPrintedRows()).EntireRow.Delete
    Loop
    
    'Prevent changes from saving
    ActiveWorkbook.Close SaveChanges:=False 'Debug.Print "Finished"
End Sub
