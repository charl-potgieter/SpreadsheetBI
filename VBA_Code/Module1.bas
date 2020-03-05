Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("B8").Select
    ActiveSheet.PivotTables("PivotTable6").PivotFields( _
        "[DummyData].[Description].[Description]").PivotItems( _
        "[DummyData].[Description].&[blah]").DrilledDown = False
    Range("B9").Select
    ActiveSheet.PivotTables("PivotTable6").PivotFields( _
        "[DummyData].[Description].[Description]").PivotItems( _
        "[DummyData].[Description].&[hello]").DrilledDown = False
End Sub
