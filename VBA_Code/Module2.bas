Attribute VB_Name = "Module2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("B8").Select
    ActiveSheet.PivotTables("PivotTable8").PivotFields( _
        "[DummyData].[Description].[Description]").LayoutBlankLine = True
End Sub
