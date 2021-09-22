Attribute VB_Name = "ZZZ_Testing"
Option Explicit



Sub TestConditionalFormatting()

    Range("J10").Select
    Selection.NumberFormat = """ERROR""; ""ERROR""; ""ERROR""; ""ERROR"""
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=WorkbookErrorStatus<>""OK"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    ExecuteExcel4Macro "(2,1,""""ERROR""; ""ERROR""; ""ERROR""; ""ERROR"""")"
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub




Sub TestConditionalFormatting2()

    Dim rng As Range
    Dim FormatConditionSheetError
    Dim FormatConditionWorkbookError As FormatCondition
    
    Set rng = ActiveSheet.Range("Heading")
    rng.FormatConditions.Delete
    
    
    Set FormatConditionSheetError = rng.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=SheetErrorStatus<>""OK""")
    
    With FormatConditionSheetError
        .NumberFormat = """SHEET ERROR""; ""SHEET ERROR"";" & _
            """SHEET ERROR""; ""SHEET ERROR"""
        .Font.Bold = True
        .Font.Italic = False
        .Font.Color = RGB(255, 0, 0)
        .Font.TintAndShade = 0
    End With
    
    Set FormatConditionWorkbookError = rng.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=WorkbookErrorStatus<>""OK""")
    
    With FormatConditionWorkbookError
        .NumberFormat = """WORKBOOK ERROR""; ""WORKBOOK ERROR"";" & _
            """WORKBOOK ERROR""; ""WORKBOOK ERROR"""
        .Font.Bold = True
        .Font.Italic = False
        .Font.Color = RGB(255, 0, 0)
        .Font.TintAndShade = 0
    End With
    
    
End Sub




