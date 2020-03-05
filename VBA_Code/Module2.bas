Attribute VB_Name = "Module2"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.PivotTables("PivotTable23").CubeFields(2).EnableMultiplePageItems _
        = True
    ActiveSheet.PivotTables("PivotTable23").PivotFields("[DateTable].[Year].[Year]" _
        ).VisibleItemsList = Array("[DateTable].[Year].&[2016]", _
        "[DateTable].[Year].&[2017]", "[DateTable].[Year].&[2019]", _
        "[DateTable].[Year].&[2020]")
    Range("E10").Select
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    With ActiveSheet.PivotTables(1).CubeFields("[DateTable].[Year]")
        .Orientation = xlPageField
        .Position = 1
    End With
End Sub
