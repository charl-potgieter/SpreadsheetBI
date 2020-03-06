Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("B8").Select
    ActiveSheet.PivotTables("PivotTable5").PivotFields( _
        "[DummyData].[Description].[Description]").LayoutSubtotalLocation = xlAtTop
End Sub
