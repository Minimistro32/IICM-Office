Attribute VB_Name = "printReccomendationSummary"
Sub printReccomendationSummary()
'
' Printcomphistory Macro
'
    Worksheets("President Sturm's").PrintOut
    Worksheets("Sister Sturm's").PrintOut Copies:=2
End Sub

