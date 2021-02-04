Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveCell.Formula2R1C1 = "={""A"",""b"",""C""}"
    Range("B15").Select
End Sub
