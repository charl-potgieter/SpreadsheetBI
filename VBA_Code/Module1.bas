Attribute VB_Name = "Module1"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.ShowAllData
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ActiveSheet.ListObjects("tbl_ReportFields").Range.AutoFilter Field:=1, _
        Criteria1:="Test Report 1"
End Sub
