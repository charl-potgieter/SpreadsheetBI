Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("C34").Select
    ActiveSheet.PivotTables("PivotTable3").HasAutoFormat = False
    Range("C32").Select
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("[Measures].[MyAmount]" _
        )
        .NumberFormat = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* ""-""??_-;_-@_-"
    End With
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("C28").Select
    With ActiveSheet.PivotTables("PivotTable3")
        .ColumnGrand = False
        .RowGrand = False
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ActiveSheet.PivotTables("PivotTable6").PivotFields( _
        "[DummyData].[Description].[Description]").VisibleItemsList = Array( _
        "[DummyData].[Description].&[hello]")
End Sub
